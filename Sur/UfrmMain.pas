unit UfrmMain;

interface

uses
  Windows, Messages, SysUtils, Classes, Controls, Forms,
  LYTray, Menus, StdCtrls, Buttons, ADODB,
  ActnList, AppEvnts, ComCtrls, ToolWin, ExtCtrls,
  registry,inifiles,Dialogs,
  StrUtils, DB,ComObj,Variants;

type
  TfrmMain = class(TForm)
    LYTray1: TLYTray;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    ADOConnection1: TADOConnection;
    ApplicationEvents1: TApplicationEvents;
    CoolBar1: TCoolBar;
    ToolBar1: TToolBar;
    ToolButton3: TToolButton;
    ToolButton4: TToolButton;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    ActionList1: TActionList;
    editpass: TAction;
    about: TAction;
    stop: TAction;
    ToolButton2: TToolButton;
    Memo1: TMemo;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    Button1: TButton;
    ToolButton5: TToolButton;
    ToolButton9: TToolButton;
    OpenDialog1: TOpenDialog;
    Timer1: TTimer;
    procedure N3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    //���Ӳ�����Ϣ���м�¼,���ظü�¼��Ψһ�����Ϊ������������
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure N1Click(Sender: TObject);
    procedure ApplicationEvents1Activate(Sender: TObject);
    procedure ToolButton7Click(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure ToolButton5Click(Sender: TObject);
  private
    { Private declarations }
    procedure WMSyscommand(var message:TWMMouse);message WM_SYSCOMMAND;
    procedure UpdateConfig;{�����ļ���Ч}
    function LoadInputPassDll:boolean;
    function MakeDBConn:boolean;
  public
    { Public declarations }
  end;

var
  frmMain: TfrmMain;

implementation

uses ucommfunction, USearchFile;

const
  CR=#$D+#$A;
  STX=#$2;ETX=#$3;ACK=#$6;NAK=#$15;
  sCryptSeed='lc';//�ӽ�������
  //SEPARATOR=#$1C;
  sCONNECTDEVELOP='����!���뿪������ϵ!' ;
  IniSection='Setup';

var
  ConnectString:string;
  GroupName:string;//
  SpecType:string ;//
  SpecStatus:string ;//
  CombinID:string;//
  LisFormCaption:string;//
  QuaContSpecNoG:string;
  QuaContSpecNo:string;
  QuaContSpecNoD:string;
  EquipChar:string;
  path_result:string;
  big_result:string;
  SdfDateFormat:string;//����ļ��������ڸ�ʽ
  ifRecLog:boolean;//�Ƿ��¼������־

//  RFM:STRING;       //��������
  hnd:integer;
  bRegister:boolean;

{$R *.dfm}

function ifRegister:boolean;
var
  HDSn,RegisterNum,EnHDSn:string;
  configini:tinifile;
  pEnHDSn:Pchar;
begin
  result:=false;
  
  HDSn:=GetHDSn('C:\')+'-'+GetHDSn('D:\')+'-'+ChangeFileExt(ExtractFileName(Application.ExeName),'');

  CONFIGINI:=TINIFILE.Create(ChangeFileExt(Application.ExeName,'.ini'));
  RegisterNum:=CONFIGINI.ReadString(IniSection,'RegisterNum','');
  CONFIGINI.Free;
  pEnHDSn:=EnCryptStr(Pchar(HDSn),sCryptSeed);
  EnHDSn:=StrPas(pEnHDSn);

  if Uppercase(EnHDSn)=Uppercase(RegisterNum) then result:=true;

  if not result then messagedlg('�Բ���,��û��ע���ע�������,��ע��!',mtinformation,[mbok],0);
end;

function GetConnectString:string;
var
  Ini:tinifile;
  userid, password, datasource, initialcatalog: string;
  ifIntegrated:boolean;//�Ƿ񼯳ɵ�¼ģʽ

  pInStr,pDeStr:Pchar;
  i:integer;
begin
  result:='';
  
  Ini := tinifile.Create(ChangeFileExt(Application.ExeName,'.INI'));
  datasource := Ini.ReadString('�������ݿ�', '������', '');
  initialcatalog := Ini.ReadString('�������ݿ�', '���ݿ�', '');
  ifIntegrated:=ini.ReadBool('�������ݿ�','���ɵ�¼ģʽ',false);
  userid := Ini.ReadString('�������ݿ�', '�û�', '');
  password := Ini.ReadString('�������ݿ�', '����', '107DFC967CDCFAAF');
  Ini.Free;
  //======����password
  pInStr:=pchar(password);
  pDeStr:=DeCryptStr(pInStr,sCryptSeed);
  setlength(password,length(pDeStr));
  for i :=1  to length(pDeStr) do password[i]:=pDeStr[i-1];
  //==========

  result := result + 'user id=' + UserID + ';';
  result := result + 'password=' + Password + ';';
  result := result + 'data source=' + datasource + ';';
  result := result + 'Initial Catalog=' + initialcatalog + ';';
  result := result + 'provider=' + 'SQLOLEDB.1' + ';';
  if ifIntegrated then
    result := result + 'Integrated Security=SSPI;';
end;

procedure TfrmMain.FormCreate(Sender: TObject);
var
  ctext        :string;
  reg          :tregistry;
begin
  ConnectString:=GetConnectString;
  
  UpdateConfig;
  if ifRegister then bRegister:=true else bRegister:=false;  

  lytray1.Hint:='���ݽ��շ���'+ExtractFileName(Application.ExeName);

//=============================��ʼ������=====================================//
    reg:=tregistry.Create;
    reg.RootKey:=HKEY_CURRENT_USER;
    reg.OpenKey('\sunyear',true);
    ctext:=reg.ReadString('pass');
    if ctext='' then
    begin
        reg:=tregistry.Create;
        reg.RootKey:=HKEY_CURRENT_USER;
        reg.OpenKey('\sunyear',true);
        reg.WriteString('pass','JIHONM{');
        //MessageBox(application.Handle,pchar('��л��ʹ�����ܼ��ϵͳ��'+chr(13)+'���ס��ʼ�����룺'+'lc'),
        //            'ϵͳ��ʾ',MB_OK+MB_ICONinformation);     //WARNING
    end;
    reg.CloseKey;
    reg.Free;
//============================================================================//
end;

procedure TfrmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    if LoadInputPassDll then action:=cafree else action:=caNone;
end;

procedure TfrmMain.N3Click(Sender: TObject);
begin
    if not LoadInputPassDll then exit;
    application.Terminate;
end;

procedure TfrmMain.N1Click(Sender: TObject);
begin
  show;
end;

procedure TfrmMain.ApplicationEvents1Activate(Sender: TObject);
begin
  hide;
end;

procedure TfrmMain.WMSyscommand(var message: TWMMouse);
begin
  inherited;
  if message.Keys=SC_MINIMIZE then hide;
  message.Result:=-1;
end;

procedure TfrmMain.ToolButton7Click(Sender: TObject);
begin
  if MakeDBConn then ConnectString:=GetConnectString;
end;

procedure TfrmMain.UpdateConfig;
var
  INI,INI_SRC_LIS:tinifile;
  autorun:boolean;
  ValueFile:string;
begin
  ini:=TINIFILE.Create(ChangeFileExt(Application.ExeName,'.ini'));

  autorun:=ini.readBool(IniSection,'�����Զ�����',false);
  ifRecLog:=ini.readBool(IniSection,'������־',false);

  ValueFile:=ini.ReadString(IniSection,'�ӿ��ļ�','');

  GroupName:=trim(ini.ReadString(IniSection,'������',''));
  EquipChar:=trim(uppercase(ini.ReadString(IniSection,'������ĸ','')));//�������Ǵ�д������һʧ��
  SpecType:=ini.ReadString(IniSection,'Ĭ����������','');
  SpecStatus:=ini.ReadString(IniSection,'Ĭ������״̬','');
  CombinID:=ini.ReadString(IniSection,'�����Ŀ����','');

  LisFormCaption:=ini.ReadString(IniSection,'����ϵͳ�������','');

  QuaContSpecNoG:=ini.ReadString(IniSection,'��ֵ�ʿ�������','9999');
  QuaContSpecNo:=ini.ReadString(IniSection,'��ֵ�ʿ�������','9998');
  QuaContSpecNoD:=ini.ReadString(IniSection,'��ֵ�ʿ�������','9997');

  SdfDateFormat:=ini.ReadString(IniSection,'����ļ��������ڸ�ʽ','YYYYMMDD');

  ini.Free;

  OperateLinkFile(application.ExeName,'\'+ChangeFileExt(ExtractFileName(Application.ExeName),'.lnk'),15,autorun);

  if FileExists(ValueFile) then
  BEGIN
    INI_SRC_LIS:=TINIFILE.Create(ValueFile);
    path_result:=INI_SRC_LIS.ReadString('LIS','PATH_RESULT','C:\');
    big_result:=INI_SRC_LIS.ReadString('LIS','big_result',',');
    INI_SRC_LIS.Free;
    Timer1.Enabled:=true;
  END else memo1.Lines.Add('û���ҵ��ļ�'+ValueFile); 
end;

function TfrmMain.LoadInputPassDll: boolean;
TYPE
    TDLLFUNC=FUNCTION:boolean;
VAR
    HLIB:THANDLE;
    DLLFUNC:TDLLFUNC;
    PassFlag:boolean;
begin
    result:=false;
    HLIB:=LOADLIBRARY('OnOffLogin.dll');
    IF HLIB=0 THEN BEGIN SHOWMESSAGE(sCONNECTDEVELOP);EXIT; END;
    DLLFUNC:=TDLLFUNC(GETPROCADDRESS(HLIB,'showfrmonofflogin'));
    IF @DLLFUNC=NIL THEN BEGIN SHOWMESSAGE(sCONNECTDEVELOP);EXIT; END;
    PassFlag:=DLLFUNC;
    FREELIBRARY(HLIB);
    result:=passflag;
end;

function TfrmMain.MakeDBConn:boolean;
var
  newconnstr,ss: string;
  Label labReadIni;
begin
  result:=false;

  labReadIni:
  newconnstr := GetConnectString;
  
  try
    ADOConnection1.Connected := false;
    ADOConnection1.ConnectionString := newconnstr;
    ADOConnection1.Connected := true;
    result:=true;
  except
  end;
  if not result then
  begin
    ss:='������'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '���ݿ�'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '���ɵ�¼ģʽ'+#2+'CheckListBox'+#2+#2+'0'+#2+#2+#3+
        '�û�'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '����'+#2+'Edit'+#2+#2+'0'+#2+#2+'1';
    if ShowOptionForm('�������ݿ�','�������ݿ�',Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
      goto labReadIni else application.Terminate;
  end;
end;

procedure TfrmMain.ToolButton2Click(Sender: TObject);
var
  ss:string;
begin
  if LoadInputPassDll then
  begin
    ss:='�ӿ��ļ�'+#2+'File'+#2+#2+'1'+#2+'ע:һ��Ϊ\Laboman4.0\lis_interface.ini'+#2+#3+
      '����ļ��������ڸ�ʽ'+#2+'Combobox'+#2+'YYYYMMDD'+#13+'YYYYMD'+#2+'0'+#2+'����2015��1��20��,YYYYMMDD->20150120,YYYYMD->2015120'+#2+#3+
      '������'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '������ĸ'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '����ϵͳ�������'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      'Ĭ����������'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      'Ĭ������״̬'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '�����Ŀ����'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '�����Զ�����'+#2+'CheckListBox'+#2+#2+'1'+#2+#2+#3+
      '������־'+#2+'CheckListBox'+#2+#2+'0'+#2+'ע:ǿ�ҽ�������������ʱ�ر�'+#2+#3+
      '��ֵ�ʿ�������'+#2+'Edit'+#2+#2+'2'+#2+#2+#3+
      '��ֵ�ʿ�������'+#2+'Edit'+#2+#2+'2'+#2+#2+#3+
      '��ֵ�ʿ�������'+#2+'Edit'+#2+#2+'2'+#2+#2;

  if ShowOptionForm('',Pchar(IniSection),Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
	  UpdateConfig;
  end;
end;

procedure TfrmMain.BitBtn2Click(Sender: TObject);
begin
  Memo1.Lines.Clear;
end;

procedure TfrmMain.BitBtn1Click(Sender: TObject);
begin
  memo1.Lines.SaveToFile('c:\comm.txt');
  showmessage('����ɹ�!');
end;

procedure TfrmMain.Button1Click(Sender: TObject);
var
  ls:Tstrings;
begin
  OpenDialog1.DefaultExt := '.txt';
  OpenDialog1.Filter := 'txt (*.txt)|*.txt';
  if not OpenDialog1.Execute then exit;
  ls:=Tstringlist.Create;
  ls.LoadFromFile(OpenDialog1.FileName);
  //ComDataPacket1Packet(nil,ls.Text);
  ls.Free;
end;

procedure TfrmMain.ToolButton5Click(Sender: TObject);
var
  ss:string;
begin
  ss:='RegisterNum'+#2+'Edit'+#2+#2+'0'+#2+'���ô���������ϵ��ַ�������������,�Ի�ȡע����'+#2;
  if bRegister then exit;
  if ShowOptionForm(Pchar('ע��:'+GetHDSn('C:\')+'-'+GetHDSn('D:\')+'-'+ChangeFileExt(ExtractFileName(Application.ExeName),'')),Pchar(IniSection),Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
    if ifRegister then bRegister:=true else bRegister:=false;
end;

procedure AFindCallBack(const filename:string;const info:tsearchrec;var quit:boolean);
var
  ls,lsValue,sList:tstrings;
  i:integer;

  SpecNo:string;
  FInts:OleVariant;
  ReceiveItemInfo:OleVariant;

  ini:Tinifile;

  //ͼ��·��
  HPLT:string;
  HRBC:string;
  HWBC:string;
  SBASO:string;
  SDIFF:string;
  SIMI:string;
  SNRBC:string;
  SPLT:string;
  SRET:string;
  SRET_E:string;
  //=========

  s1:string;
  i0:TDateTime;//�ϴμ���ʱ��
  i1:TDateTime;//���μ���ʱ��
  sName:string;//�ļ���
  fs:TFormatSettings;
  s2:string;
begin
  sName:=ExtractFileName(filename);
  
  sList:=TStringList.Create;
  ExtractStrings(['_'],[],PChar(sName),sList);
  if sList.Count<2 then begin sList.Free;exit;end;
  s1:=sList[0]+'_'+sList[1];
  sList.Free;
    
  ini:=TINIFILE.Create(ChangeFileExt(Application.ExeName,'.ini'));
  i0:=ini.ReadDateTime(FormatDateTime('YYYYMMDD',now),s1,0);
  ini.Free;

  ls:=Tstringlist.Create;
  ls.LoadFromFile(filename);
  if ls.Count<=0 then begin ls.Free;exit;end;//���������û��cdf�ļ���д�꣬��ȴ�д��

  //���μ���ʱ��
  i1:=1;
  for i :=0  to ls.Count-1 do
  begin
    lsValue:=StrToList(ls[i],big_result);//��ÿ�е��뵽�ַ����б���

    if lsValue.Count<20 then begin lsValue.Free;continue;end;
    s2:=StringReplace(lsValue[19],'/','-',[rfReplaceAll, rfIgnoreCase]);

    if lsValue[0]<>'00' then begin lsValue.Free;continue;end;

    fs.DateSeparator:='-';
    fs.TimeSeparator:=':';
    fs.ShortDateFormat:='YYYY-MM-DD hh:nn:ss';
    i1:=StrtoDateTimeDef(s2,i1,fs);

    lsValue.Free;
  end;
  //==========

  if i1<=i0 then begin ls.Free;exit;end;//���ļ��Ѿ���������Ѵ������ǰ����
  
  ini:=TINIFILE.Create(ChangeFileExt(Application.ExeName,'.ini'));
  ini.WriteDateTime(FormatDateTime('YYYYMMDD',now),s1,i1);
  ini.Free;

  if length(frmMain.memo1.Lines.Text)>=60000 then frmMain.memo1.Lines.Clear;//memoֻ�ܽ���64K���ַ�
  frmMain.memo1.Lines.Add(filename);

  //ȡͼ������
  for i :=0  to ls.Count-1 do
  begin
    lsValue:=StrToList(ls[i],big_result);//��ÿ�е��뵽�ַ����б���

    if lsValue.Count<4 then continue;

    if uppercase(lsValue[2])='HPLT' then HPLT:=lsValue[3];
    if uppercase(lsValue[2])='HRBC' then HRBC:=lsValue[3];
    if uppercase(lsValue[2])='HWBC' then HWBC:=lsValue[3];
    if uppercase(lsValue[2])='SBASO' then SBASO:=lsValue[3];
    if uppercase(lsValue[2])='SDIFF' then SDIFF:=lsValue[3];
    if uppercase(lsValue[2])='SIMI' then SIMI:=lsValue[3];
    if uppercase(lsValue[2])='SNRBC' then SNRBC:=lsValue[3];
    if uppercase(lsValue[2])='SPLT' then SPLT:=lsValue[3];
    if uppercase(lsValue[2])='SRET' then SRET:=lsValue[3];
    if uppercase(lsValue[2])='SRET-E' then SRET_E:=lsValue[3];

    lsValue.Free;
  end;
  //============

  ReceiveItemInfo:=VarArrayCreate([0,ls.Count-1],varVariant);

  for i :=0  to ls.Count-1 do
  begin
    lsValue:=StrToList(ls[i],big_result);//��ÿ�е��뵽�ַ����б���

    if lsValue.Count<4 then
    begin
      ReceiveItemInfo[i]:=VarArrayof(['','','','']);
      continue;
    end;

    if lsValue[0]='0' then SpecNo:=rightstr('0000'+lsValue[3],4);

    if uppercase(lsValue[1])='PLT' then ReceiveItemInfo[i]:=VarArrayof([lsValue[1],lsValue[3],'',HPLT])
    else if uppercase(lsValue[1])='RBC' then ReceiveItemInfo[i]:=VarArrayof([lsValue[1],lsValue[3],'',HRBC])
    else if uppercase(lsValue[1])='WBC' then ReceiveItemInfo[i]:=VarArrayof([lsValue[1],lsValue[3],'',HWBC])
    else if uppercase(lsValue[1])='BASO#' then ReceiveItemInfo[i]:=VarArrayof([lsValue[1],lsValue[3],'',SBASO])
    else if uppercase(lsValue[1])='MPV' then ReceiveItemInfo[i]:=VarArrayof([lsValue[1],lsValue[3],'',SDIFF])
    else if uppercase(lsValue[1])='MONO#' then ReceiveItemInfo[i]:=VarArrayof([lsValue[1],lsValue[3],'',SIMI])
    else if uppercase(lsValue[1])='NRBC#' then ReceiveItemInfo[i]:=VarArrayof([lsValue[1],lsValue[3],'',SNRBC])
    else if uppercase(lsValue[1])='HCT' then ReceiveItemInfo[i]:=VarArrayof([lsValue[1],lsValue[3],'',SPLT])
    else if uppercase(lsValue[1])='RET#' then ReceiveItemInfo[i]:=VarArrayof([lsValue[1],lsValue[3],'',SRET])
    else if uppercase(lsValue[1])='RET%' then ReceiveItemInfo[i]:=VarArrayof([lsValue[1],lsValue[3],'',SRET_E])
    else ReceiveItemInfo[i]:=VarArrayof([lsValue[1],lsValue[3],'','']);

    lsValue.Free;
  end;
  
  ls.Free;

  if bRegister then
  begin
    FInts :=CreateOleObject('Data2LisSvr.Data2Lis');
    FInts.fData2Lis(ReceiveItemInfo,(SpecNo),'',
      (GroupName),(SpecType),(SpecStatus),(EquipChar),
      (CombinID),'',(LisFormCaption),(ConnectString),
      (QuaContSpecNoG),(QuaContSpecNo),(QuaContSpecNoD),'',
      ifRecLog,true,'����');
    if not VarIsEmpty(FInts) then FInts:= unAssigned;
  end;
end;

procedure TfrmMain.Timer1Timer(Sender: TObject);
var
  qqq:boolean;

  lsSection    :TStrings;
  ini          :Tinifile;
  i            :integer;
begin
  (Sender as TTimer).Enabled:=false;

  lsSection:=Tstringlist.Create;
  ini:=TINIFILE.Create(ChangeFileExt(Application.ExeName,'.ini'));
  ini.ReadSections(lsSection);
  for i :=0  to lsSection.Count-1 do
  begin
    if (leftstr(lsSection[i],2)='20')and(lsSection[i]<>FormatDateTime('YYYYMMDD',now)) then ini.EraseSection(lsSection[i]);
  end;
  ini.Free;
  lsSection.Free;

  qqq:=false;
  findfile(qqq,PATH_RESULT,FormatDateTime(SdfDateFormat,now)+'_*.cdf',AFindCallBack,false,true);

  (Sender as TTimer).Enabled:=true;
end;

initialization
    hnd := CreateMutex(nil, True, Pchar(ExtractFileName(Application.ExeName)));
    if GetLastError = ERROR_ALREADY_EXISTS then
    begin
        MessageBox(application.Handle,pchar('�ó������������У�'),
                    'ϵͳ��ʾ',MB_OK+MB_ICONinformation);   
        Halt;
    end;

finalization
    if hnd <> 0 then CloseHandle(hnd);

end.
