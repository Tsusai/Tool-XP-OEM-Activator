unit Main;

interface

uses
	Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
	Dialogs, StdCtrls, ComCtrls, Buttons;

type
	TForm1 = class(TForm)
		OSBox: TComboBox;
		BiosBox: TComboBox;
		MessageBox: TMemo;
		Label1: TLabel;
		ActivateButton: TBitBtn;
		procedure OSBoxSelect(Sender: TObject);
		procedure FormCreate(Sender: TObject);
		procedure FormDestroy(Sender: TObject);
		procedure FormShow(Sender: TObject);
		procedure ActivateButtonClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
	private
		{ Private declarations }
	public
		{ Public declarations }
	end;

type
	TDataList = class(TStringList)
	Public
		Destructor Destroy; override;
	end;

type
	TOEM = class
	public
		XPHome : string;
		XPPro : string;
		XPMCE : string;
		XPTab : string;
	end;


var
	Form1: TForm1;

	AppPath : string;

	XPCertList : TStringList;

	DataList : TDataList;

implementation
uses WindowsAPI, IniFiles, StrUtils, magwmi, ShellAPI, MMSystem, Registry;

{$R *.dfm}
//{$R Activator.REC}

function LoadFolders : boolean; forward;
function LoadKeys : boolean; forward;


procedure AutoSetup; forward;

procedure MsgBoxAdd(Line : string);
begin
	Form1.MessageBox.Lines.Add(Line);
	SendMessage(Form1.MessageBox.Handle, EM_SCROLLCARET, 0, 0);
end;

{function IsLike(AString, Pattern: string): boolean;
var
	 j, n, n1, n2: integer ;
	 p1, p2: pchar ;
label
	match, nomatch;
begin
	AString := UpperCase(AString) ;
	Pattern := UpperCase(Pattern) ;
	n1 := Length(AString) ;
	n2 := Length(Pattern) ;
	if n1 < n2 then n := n1 else n := n2;
	p1 := pchar(AString) ;
	p2 := pchar(Pattern) ;
	for j := 1 to n do begin
		if p2^ = '*' then goto match;
		if (p2^ <> '?') and ( p2^ <> p1^ ) then goto nomatch;
		inc(p1) ; inc(p2) ;
	end;
	if n1 > n2 then
	begin
		nomatch:
			Result := False;
			exit;
	end else if n1 < n2 then
	begin
		for j := n1 + 1 to n2 do
		begin
			if not ( p2^ in ['*','?'] ) then goto nomatch ;
			inc(p2) ;
		 end;
	end;
	match:
	Result := True
end;   }

function FixUpperCase(Input : string) : string;
begin
	Result := LowerCase(Input);
	Result[1] := UpCase(Result[1]);
end;

function InitGlobals : boolean;
begin
	AppPath := ExtractFilePath(ParamStr(0));
	try
		DataList := TDataList.Create;
		XPCertList := TStringList.Create;
		XPCertList.Sorted := True;

		LoadFolders;
		LoadKeys;
		Result := true;
	finally
	end;
end;

procedure DestroyGlobals;
begin
	XPCertList.Free;
	DataList.Free;
end;

function LoadKeys : boolean;
var
	idx : integer;
	Keyini : TMemIniFile;
	KeyData : TOEM;
begin
	Result := true;
	DataList.Clear;

	with TStringList.Create do
	begin
		try
			Duplicates := dupIgnore;
			for idx := 0 to XPCertList.Count - 1 do Add(XPCertList[idx]);


			for idx := 0 to Count - 1 do
			begin
				Keyini:= TMemIniFile.Create(AppPath+'OEMFiles\'+Strings[idx]+'\Keys.ini');
				try
					KeyData := TOEM.Create;

					KeyData.XPHome := Keyini.ReadString('Keys','XP Home','');
					KeyData.XPPro := Keyini.ReadString('Keys','XP Professional','');
					KeyData.XPMCE := Keyini.ReadString('Keys','XP Media Center','');
					KeyData.XPTab := Keyini.ReadString('Keys','XP Tablet','');

					DataList.AddObject(Strings[idx],KeyData);
				finally
					Keyini.Free;
				end;
			end;
		finally
			Free;
		end;
	end;
end;

function LoadFolders : boolean;
var
	searchResult : TSearchRec;
	idx : integer;
begin
	Result := true;
	XPCertList.Clear;
	if FindFirst(AppPath+'OEMFiles\*.*', faDirectory, searchResult) = 0 then
	begin
		repeat
			IF (searchResult.Attr AND faDirectory > 0) AND
			(searchResult.Name <> '.') AND
			(searchResult.Name <> '..') then
			begin
				XPCertList.Add(searchResult.Name);
			end;
		until FindNext(searchResult) <> 0;
		FindClose(searchResult);

		for Idx := XPCertList.Count - 1 downto 0 do
		begin
			if NOT (FileExists(AppPath+'OEMFiles\'+XPCertList[idx]+'\oembios.bin') and
				FileExists(AppPath+'OEMFiles\'+XPCertList[idx]+'\oembios.dat') and
				FileExists(AppPath+'OEMFiles\'+XPCertList[idx]+'\oembios.sig') and
				FileExists(AppPath+'OEMFiles\'+XPCertList[idx]+'\oembios.cat') and
				FileExists(AppPath+'OEMFiles\'+XPCertList[idx]+'\Keys.ini')) then
			begin
				MsgBoxAdd('ERROR: Missing OEMBIOS files or Keys missing for ' + XPCertList[idx] + '. Removed from XP List');
				XPCertList.Delete(idx);
			end;
		end;

	end;
end;

function ExpandEnvironment(const strValue: string): string;
var
	chrResult: array[0..1023] of Char;
	wrdReturn: DWORD;
begin
	wrdReturn := ExpandEnvironmentStrings(PChar(strValue), chrResult, 1024);
	if wrdReturn = 0 then
		Result := strValue
	else
	begin
		Result := Trim(chrResult);
	end;
end;


function FindOS : boolean;
var
	OS: TWinVersion;
begin

	Result := true;
	OS := GetWindowsVersion;
	case OS of
		wvWinXPHome: Form1.OSBox.ItemIndex := 0;
		wvWinXPPro: Form1.OSBox.ItemIndex := 1;
		wvWinXPMCE: Form1.OSBox.ItemIndex := 2;
		wvWinXPTab: Form1.OSBox.ItemIndex := 3;
		wvWinVistaHB: Form1.OSBox.ItemIndex := 4;
		wvWinVistaHP: Form1.OSBox.ItemIndex := 5;
		wvWinVistaBS: Form1.OSBox.ItemIndex := 6;
		wvWinVistaUT: Form1.OSBox.ItemIndex := 7;
		wvWinSevenHP: Form1.OSBox.ItemIndex := 8;
		wvWinSevenPR: Form1.OSBox.ItemIndex := 9;
		wvWinSevenUT: Form1.OSBox.ItemIndex := 10;
	else
		begin
			ShowMessage('Unknown OS');
			Result := false;
			Exit;
		end;
	end;

	if IsSafeMode then
	begin
		ShowMessage('Safe Mode Detected.  Please run in Normal Mode');
		Result := false;
		Exit;
	end;

	if Result then Form1.OSBoxSelect(Form1);
end;

procedure FindOEM;
var
	SysInfo : string;
	idx : integer;
begin
	MsgBoxAdd('MB: ' + MagWmiGetBaseBoard);
	MsgBoxAdd('Bios: ' + MagWmiGetSMBIOS);
	if Form1.BiosBox.Items.Count > 0 then
	Begin
		Form1.BiosBox.ItemIndex := -1;
		SysInfo := MagWmiGetBaseBoard + ' ' + MagWmiGetSMBIOS;
		for idx := 0 to Form1.BiosBox.Items.Count - 1 do
		begin
			if AnsiContainsStr(LowerCase(SysInfo),LowerCase(Form1.BiosBox.Items[idx])) then
			begin
				Form1.BiosBox.ItemIndex := idx;
				Exit;
			end;
		end;
	end;
end;

procedure AutoSetup;
begin
	FindOS; //calls findoem
end;

procedure GetOEMData(
	var Key : string;
	var Cert : string
);
var
	idx : integer;
	tmp : string;
begin
	Key := '';
	Cert := '';
	tmp := Form1.BiosBox.Items[Form1.BiosBox.ItemIndex];
	idx := DataList.IndexOf(Form1.BiosBox.Items[Form1.BiosBox.ItemIndex]);
	if idx > -1 then
	begin
		case Form1.OSBox.ItemIndex of
		0: Key := TOEM(DataList.Objects[idx]).XPHome;
		1: Key := TOEM(DataList.Objects[idx]).XPPro;
		2: Key := TOEM(DataList.Objects[idx]).XPMCE;
		3: Key := TOEM(DataList.Objects[idx]).XPTab;
		end;
	end;
end;

procedure TForm1.ActivateButtonClick(Sender: TObject);
var
	Key : string;
	Cert : string;
	Path : string;
	MB : string;
	Temp : string;
begin
	//Activate commands here.
	MsgBoxAdd('Getting key');
	GetOEMData(Key,Cert);
	if Key = '' then
	begin
		MsgBoxAdd('Couldn''t activate.  No key available for this version');
		exit;
	end;
	{if Key = 'QUERY' then
	begin
		Key := InputBox('Key Entry','Please enter original key'+#10#13+
			'(XXXXX-XXXXX-XXXXX-XXXXX-XXXXX','');
	end;}
	MsgBoxAdd('setting path');
	MB := AppPath + 'OEMFiles\' + BiosBox.Items[BiosBox.ItemIndex] + '\';
	MsgBoxAdd('figuring out what version');
	case Form1.OSBox.ItemIndex of
		0..3:
		begin
			//Check and copy files
			Temp := ExpandEnvironment('%TEMP%\');
			Path := ExpandEnvironment('%Systemroot%\System32\');

			if FileExists(Path+'dllcache\oembios.bin') and
				FileExists(Path+'dllcache\oembios.cat') and
				FileExists(Path+'dllcache\oembios.dat') and
				FileExists(Path+'dllcache\oembios.sig') then
			begin
				MsgBoxAdd('Removing DLL Cache Backup');
				scShellDeleteFile(Path+'dllcache\oembios.*',true,false,false);
			end;

			MsgBoxAdd('Getting permission from SFC');

			if IsSFCProtected(Path+'oembios.bin') then SetupSFCException(Path+'oembios.bin');
			if IsSFCProtected(Path+'oembios.dat') then SetupSFCException(Path+'oembios.dat');
			if IsSFCProtected(Path+'oembios.sig') then SetupSFCException(Path+'oembios.sig');
			if IsSFCProtected(Path+'CatRoot\{F750E6C3-38EE-11D1-85E5-00C04FC295EE}\oembios.cat')
			 then SetupSFCException(Path+'CatRoot\{F750E6C3-38EE-11D1-85E5-00C04FC295EE}\oembios.cat');

			MsgBoxAdd('Moving Old OEMBIOS files to ' + Temp);

			scShellMoveFile(Path+'oembios.*', Temp,true);
			scShellMoveFile(Path+'CatRoot\{F750E6C3-38EE-11D1-85E5-00C04FC295EE}\oembios.cat', Temp, true);

			MsgBoxAdd('Applying new OEMBIOS files');
			//Where they need to go!
			scShellCopyFile(MB+'oembios.bin', Path, true);
			scShellCopyFile(MB+'oembios.dat', Path, true);
			scShellCopyFile(MB+'oembios.sig', Path, true);
			scShellCopyFile(MB+'oembios.cat', Path+ 'CatRoot\{F750E6C3-38EE-11D1-85E5-00C04FC295EE}\', true);

			//DLLCACHE
			scShellCopyFile(MB+'oembios.bin', Path+'\dllcache', true);
			scShellCopyFile(MB+'oembios.dat', Path+'\dllcache', true);
			scShellCopyFile(MB+'oembios.sig', Path+'\dllcache', true);
			scShellCopyFile(MB+'oembios.cat', Path+'\dllcache', true);

			MsgBoxAdd('Changing Key');
			RunAndWait('ChangeKey.vbs', Key);

			MsgBoxAdd('Removing Start Menu Shortcut');

			Temp := ExpandEnvironment('%ALLUSERSPROFILE%\Start Menu\');
			if FileExists(Temp + 'Activate Windows.lnk') then
			begin
				scShellDeleteFile(Temp + 'Activate Windows.lnk', true,false,false);
			end;

			with TRegistry.Create do
			try
				RootKey := HKEY_LOCAL_MACHINE;
				if OpenKey('\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce',false) then
				begin
					WriteString('LoadAct','"'+Path+'oobe\msoobe.exe" /A');
				end;
			finally
				Free;
			end;


			MsgBoxAdd('Done');
			if (MessageDlg('System Must reboot for changes to take effect.  Reboot now?',
				mtConfirmation, [mbYes,mbNo], 0) = Integer(mrYes)) then
			begin
				ShellExecute(GetDesktopWindow,'open','shutdown','-r -t 5 -c "System will'+
				' reboot to apply changes',nil, SW_SHOWNORMAL);
				Application.Terminate;
			end;
			{}
		end;
		4..10:
		begin
			if Cert <> '' then
			begin
				MsgBoxAdd('Installing Certificate: ' + MB + Cert);
				RunAndWait('slmgr.vbs','-ilc ' + MB + Cert);

				MsgBoxAdd('Installing Key: ' + Key);
				RunAndWait('slmgr.vbs','-ipk ' + Key);

				MsgBoxAdd('Activating');
				RunAndWait('slmgr.vbs','-ato');

				MsgBoxAdd('Done');
			end else
			begin
				MsgBoxAdd('Couldn''t activate.  No Certificate found');
				exit;
			end;
		end;
	end;
end;

{procedure TForm1.Button1Click(Sender: TObject);
var
	Key : string;

const Path = 'c:\windows\system32\';

begin
Key := InputBox('Key Entry','Please enter original key'+#10#13+
			'(XXXXX-XXXXX-XXXXX-XXXXX-XXXXX','');
if IsLike(Key,'?????-?????-?????-?????-?????') then
MsgBoxAdd('Valid key') else MsgBoxAdd('Not valid key');

end;}

procedure TForm1.FormActivate(Sender: TObject);
begin
	SendMessage(MessageBox.Handle, EM_SCROLLCARET, 0, 0);
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
	InitGlobals;
	mciSendString('Set cdaudio door open wait', nil, 0, 0);
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
	DestroyGlobals;
end;

procedure TForm1.FormShow(Sender: TObject);
var ALeft, ATop: integer;
begin
	//Center
	ALeft := (Screen.Width - Width) div 2;
	ATop  := (Screen.Height - Height) div 2;
	{ prevents form being twice repainted! }
	Self.SetBounds(ALeft, ATop, Width, Height);

	//Find whatever we're using
	AutoSetup;
end;

procedure TForm1.OSBoxSelect(Sender: TObject);
begin
	case OSBox.ItemIndex of
	0..3:
		begin
			BiosBox.Items.Assign(XPCertList);
			FindOEM;
		end;

	else BiosBox.Items.Clear;
	end;
end;

destructor TDataList.Destroy;
var
	Index : Integer;
begin
	if (Count >0) then
	begin
		for Index := Count -1 downto 0 do
		begin
			Self.Objects[Index].Free;
			Self.Delete(Index);
		end;
	end;
	inherited;
end;

end.

