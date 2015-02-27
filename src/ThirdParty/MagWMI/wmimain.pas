unit wmimain;

{$WARN UNSAFE_TYPE off}
{$WARN UNSAFE_CAST off}
{$WARN UNSAFE_CODE off}
{$WARN SYMBOL_PLATFORM OFF}
{$WARN SYMBOL_LIBRARY OFF}
{$WARN SYMBOL_DEPRECATED OFF}

{
Magenta Systems WMI and SMART Component demo application v5.2
Updated by Angus Robertson, Magenta Systems Ltd, England, 5th March 2009
delphi@magsys.co.uk, http://www.magsys.co.uk/delphi/
Copyright 2009, Magenta Systems Ltd

10th January 2004 - Release 4.93
12th July 2004 - Release 4.94 - added disk Smart failure info
14th October 2004 - Release 4.95 - added SCSI disk serial checking
9th January 2005  - Release 4.96 - added MagWmiCloseWin to close down windows
22nd October 2005 - Release 5.0  - separated from TMagRas
29th July 2008    - Release 5.1  - compability with unicode in Delphi 2009
5th March 2009    - Release 5.2  - fixed memory leaks with OleVariants
                                   better error handling



}

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, Spin, ExtCtrls, magwmi, magsubs1 ;

type
  TForm1 = class(TForm)
    edtClass: TComboBox;
    edtComputer: TEdit;
    edtUser: TEdit;
    edtPass: TEdit;
    Label4: TLabel;
    Label7: TLabel;
    Label5: TLabel;
    Label8: TLabel;
    ListView: TListView;
    doGetClass: TButton;
    doExit: TButton;
    doMB: TButton;
    ResInfo: TEdit;
    doBIOS: TButton;
    Label1: TLabel;
    doDiskModel: TButton;
    DiskNum: TSpinEdit;
    doDiskSerial: TButton;
    doBootTime: TButton;
    Label2: TLabel;
    doCommand: TButton;
    OneCommand: TComboBox;
    Label3: TLabel;
    OneProp: TComboBox;
    Panel1: TPanel;
    doIPAddr: TButton;
    SubNetMask: TEdit;
    Label6: TLabel;
    Label10: TLabel;
    IPAddress: TEdit;
    IPGateway: TEdit;
    Label9: TLabel;
    edtNamespace: TComboBox;
    doRenameComp: TButton;
    NewCompName: TEdit;
    Label11: TLabel;
    doSmart: TButton;
    doReboot: TButton;
    doCloseDown: TButton;
    doMemory: TButton;
    StatusBar: TStatusBar;
    procedure doGetClassClick(Sender: TObject);
    procedure doExitClick(Sender: TObject);
    procedure doMBClick(Sender: TObject);
    procedure doBIOSClick(Sender: TObject);
    procedure doCommandClick(Sender: TObject);
    procedure doBootTimeClick(Sender: TObject);
    procedure doDiskSerialClick(Sender: TObject);
    procedure doDiskModelClick(Sender: TObject);
    procedure doIPAddrClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure doRenameCompClick(Sender: TObject);
    procedure doSmartClick(Sender: TObject);
    procedure doRebootClick(Sender: TObject);
    procedure doCloseDownClick(Sender: TObject);
    procedure doMemoryClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}


procedure TForm1.doGetClassClick(Sender: TObject);
var
    rows, instances, I, J: integer ;
    WmiResults: T2DimStrArray ;
    OldCursor: TCursor ;
    errstr: string ;
begin
    doGetClass.Enabled := false ;
    OldCursor := Screen.Cursor ;
    Screen.Cursor := crHourGlass ;
    StatusBar.SimpleText := '' ;
    try
        ListView.Items.Clear ;
        Application.ProcessMessages ;
        rows := MagWmiGetInfoEx (edtComputer.Text, edtNameSpace.Text, edtUser.Text,
                           edtPass.Text, edtClass.Text, WmiResults, instances, errstr) ;
        if rows > 0 then
        begin
            StatusBar.SimpleText := 'Instances: ' + IntToStr (instances) ;
            if instances >= ListView.Columns.Count then
                                        instances := ListView.Columns.Count - 1 ;
            for J := 0 to instances do
                            ListView.Columns.Items [J].Caption := WmiResults [J, 0] ;
            for I := 1 to rows do
            begin
                with ListView.Items.Add do
                begin
                    Caption := WmiResults [0, I] ;
                    for J := 1 to instances do
                        SubItems.Add (WmiResults [J, I]) ;
                end ;
            end ;
        end
        else if rows = -1 then
            StatusBar.SimpleText := 'Error: ' + errstr
        else
           StatusBar.SimpleText := 'Instances: None' ;
    finally
        doGetClass.Enabled := true ;
        Screen.Cursor := OldCursor ; 
        WmiResults := Nil ;
    end ;
end;

procedure TForm1.doExitClick(Sender: TObject);
begin
    Close ;
end;

procedure TForm1.doMBClick(Sender: TObject);
begin
    ResInfo.Text := MagWmiGetBaseBoard ;
end;

procedure TForm1.doBIOSClick(Sender: TObject);
begin
    ResInfo.Text := MagWmiGetSMBIOS ;
end;

procedure TForm1.doCommandClick(Sender: TObject);
var
    res: string ;
begin
    doCommand.Enabled := false ;
    ResInfo.Text := '' ;
    try
        if MagWmiGetOneQ (OneCommand.Text, OneProp.Text, res) > 0 then
                                                          ResInfo.Text := res ;
    finally
        doCommand.Enabled := true ;
    end ;
end;

procedure TForm1.doBootTimeClick(Sender: TObject);
var
    When: TDateTime ;
begin
    ResInfo.Text := '' ;
    When := MagWmiGetLastBootDT ;
    if When > 10 then ResInfo.Text := DateTimeToStr (When) ;
end;

procedure TForm1.doDiskSerialClick(Sender: TObject);
begin
    ResInfo.Text := MagWmiGetDiskSerial (DiskNum.Value) ;
end ;

procedure TForm1.doDiskModelClick(Sender: TObject);
begin
    ResInfo.Text := MagWmiGetDiskModel (DiskNum.Value) ;
end;

procedure TForm1.doIPAddrClick(Sender: TObject);
var
    adapter: string ;
    res, index: integer ;
    IPAddresses, SubnetMasks, IPGateways: StringArray;
    GatewayCosts: TIntegerArray ;
begin
    ResInfo.Text := '' ;
    adapter := '' ;
    SetLength (IPAddresses, 1) ;  // note, may be more than one address/mask
    SetLength (SubnetMasks, 1) ;
    SetLength (IPGateways, 1) ;
    SetLength (GatewayCosts, 1) ;
    IPAddresses [0] := IPAddress.Text ;
    SubnetMasks [0] := SubnetMask.Text ;
    IPGateways [0] := IPGateway.Text ;
    GatewayCosts [0] := 10 ;
    index := MagWmiFindAdaptor (adapter) ;  // looks for current Local Areas adaptor
    if index < 0 then
    begin
        ResInfo.Text := 'Can Not Find Single Adapter' ;
        exit ;
    end ;
    res := MagWmiNetSetIPAddr (index, IPAddresses, SubnetMasks) ;
    ResInfo.Text := adapter +  ' - Change IP Result: ' + IntToStr (res) ;
    if (res < 0) or (res > 1) then exit ;
    Res := MagWmiNetSetGateway (index, IPGateways, GatewayCosts) ;
    ResInfo.Text := adapter +  ' - Change IP Result: ' + IntToStr (res) ;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
    NewCompName.Text := GetCompName ;
end;

procedure TForm1.doRenameCompClick(Sender: TObject);
var
    res: integer ;
begin
    if NewCompName.Text = GetCompName then exit ;
    if NewCompName.Text = '' then exit ;
    res := MagWmiRenameComp (NewCompName.Text, edtUser.Text, edtPass.Text) ;
    if res = 0 then
        ResInfo.Text := 'Rename Computer OK, Must Reboot Now'
    else
        ResInfo.Text := 'Rename Computer Failed: ' + IntToStr (res) ;
    end;

procedure TForm1.doSmartClick(Sender: TObject);
var
    errinfo, model, serial: string ;
    diskbytes: int64 ;
    SmartResult: TSmartResult ;
    I: integer ;

    function GetYN (value: boolean): string ;
    begin
        result := 'No' ;
        if value then result := 'Yes' ;
    end ;

begin
    ListView.Items.Clear ;
    StatusBar.SimpleText := '' ;
    if NOT MagWmiSmartDiskInfo (DiskNum.Value, errinfo, model, serial, diskbytes) then
    begin
        if MagWmiScsiDiskInfo (DiskNum.Value, errinfo, model, serial, diskbytes) then
            ResInfo.Text := 'SCSI Serial: ' + serial {+
                                             ', Size = ' + Int64ToCStr (diskbytes) }
        else
            ResInfo.Text := errinfo ;
        exit ;
    end ;
    ResInfo.Text := 'IDE Model: ' + model + ', Serial: ' + serial +
                                             ', Size = ' + Int64ToCStr (diskbytes) ;
    if NOT MagWmiSmartDiskFail (DiskNum.Value, SmartResult, errinfo) then
    begin
        ResInfo.Text := errinfo ;
        exit ;
    end ;
    if SmartResult.TotalAttrs = 0 then
    begin
        ResInfo.Text := ResInfo.Text + ' No SMART Attributes Returned - ' + errinfo ;
        exit ;
    end ;
    StatusBar.SimpleText := 'SMART Attributes'  ;
    with ListView.Columns do
    begin
        Items [0].Caption := 'Attribute' ;
        Items [1].Caption := 'Name' ; 
        Items [2].Caption := 'State' ;
        Items [3].Caption := 'Current' ;
        Items [4].Caption := 'Worst' ;
        Items [5].Caption := 'Threshold' ;
        Items [6].Caption := 'Raw Value' ;
        Items [7].Caption := 'Pre-Fail?' ;
        Items [8].Caption := 'Events?' ;
        Items [9].Caption := 'Error Rate?' ;
    end ;
    ListView.Items.Add.Caption := 'Temp ' + IntToStr (SmartResult.Temperature) + '�C' ;
//    ListView.Items.Add.Caption := 'Hours ' + IntToStr (SmartResult.HoursRunning) ;
    ListView.Items.Add.Caption := 'Realloc Sec ' + IntToStr (SmartResult.ReallocSector) ;
    if SmartResult.SmartFailTot <> 0 then
        ListView.Items.Add.Caption := 'SMART Test Failed, Bad Attributes ' +
                                             IntToStr (SmartResult.SmartFailTot) 
    else
        ListView.Items.Add.Caption := 'SMART Test Passed' ;
    for I := 0 to Pred (SmartResult.TotalAttrs) do
    begin
        with ListView.Items.Add do
        begin
            Caption := IntToStr (SmartResult.AttrNum [I]) ;
            SubItems.Add (SmartResult.AttrName [I]) ;
            SubItems.Add (SmartResult.AttrState [I]) ;
            SubItems.Add (IntToStr (SmartResult.AttrCurValue [I])) ;
            SubItems.Add (IntToStr (SmartResult.AttrWorstVal [I])) ;
            SubItems.Add (IntToStr (SmartResult.AttrThreshold [I])) ;
            SubItems.Add (IntToCStr (SmartResult.AttrRawValue [I])) ;
            SubItems.Add (GetYN (SmartResult.AttrPreFail [I])) ;
            SubItems.Add (GetYN (SmartResult.AttrEvents [I])) ;
            SubItems.Add (GetYN (SmartResult.AttrErrorRate [I])) ;
        end ;
    end ;
end;

procedure TForm1.doRebootClick(Sender: TObject);
var
    S: string ;
begin
    ResInfo.Text := '' ;
    if edtComputer.Text = '.' then edtComputer.Text := GetCompName ;
    if Application.MessageBox (PChar ('Confirm Reboot PC ' + edtComputer.Text + ' Now'),
                             'WMI - Reboot PC', MB_OKCANCEL) <> IDOK then exit  ;
    if MagWmiCloseWin (edtComputer.Text, edtUser.Text, edtPass.Text,true, S) = 0 then
                                                                S := 'PC Reboot Accepted' ;
    ResInfo.Text := S ;
end;

procedure TForm1.doCloseDownClick(Sender: TObject);
var
    S: string ;
begin
    ResInfo.Text := '' ;
    if edtComputer.Text = '.' then edtComputer.Text := GetCompName ;
    if Application.MessageBox (Pchar ('Confirm Power Down PC ' + edtComputer.Text + ' Now'),
                             'WMI - Power Down PC', MB_OKCANCEL) <> IDOK then exit  ;
    if MagWmiCloseWin (edtComputer.Text, edtUser.Text, edtPass.Text,false, S) = 0 then
                                                         S := 'PC Power Down Accepted' ;
    ResInfo.Text := S ;
end;

procedure TForm1.doMemoryClick(Sender: TObject);
var
    WmiMemoryRec: TWmiMemoryRec ;
begin
    ResInfo.Text := '' ; 
    WmiMemoryRec := MagWmiGetMemory ;
    with WmiMemoryRec do
    begin
        ResInfo.Text := 'FreePhysicalMemory: ' + IntToKbyte (FreePhysicalMemory) +
            ', FreeSpaceInPagingFiles: ' + IntToKbyte (FreeSpaceInPagingFiles) +
            ', FreeVirtualMemory: ' + IntToKbyte (FreeVirtualMemory) ; 
      { SizeStoredInPagingFiles
        TotalSwapSpaceSize
        TotalVirtualMemorySize
        TotalVisibleMemorySize  }
    end ;
end;

end.
