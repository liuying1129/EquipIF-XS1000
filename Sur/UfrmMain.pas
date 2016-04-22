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
    //增加病人信息表中记录,返回该记录的唯一编号作为检验结果表的外键
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
    procedure UpdateConfig;{配置文件生效}
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
  sCryptSeed='lc';//加解密种子
  //SEPARATOR=#$1C;
  sCONNECTDEVELOP='错误!请与开发商联系!' ;
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
  SdfDateFormat:string;//结果文件名的日期格式
  ifRecLog:boolean;//是否记录调试日志

//  RFM:STRING;       //返回数据
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

  if not result then messagedlg('对不起,您没有注册或注册码错误,请注册!',mtinformation,[mbok],0);
end;

function GetConnectString:string;
var
  Ini:tinifile;
  userid, password, datasource, initialcatalog: string;
  ifIntegrated:boolean;//是否集成登录模式

  pInStr,pDeStr:Pchar;
  i:integer;
begin
  result:='';
  
  Ini := tinifile.Create(ChangeFileExt(Application.ExeName,'.INI'));
  datasource := Ini.ReadString('连接数据库', '服务器', '');
  initialcatalog := Ini.ReadString('连接数据库', '数据库', '');
  ifIntegrated:=ini.ReadBool('连接数据库','集成登录模式',false);
  userid := Ini.ReadString('连接数据库', '用户', '');
  password := Ini.ReadString('连接数据库', '口令', '107DFC967CDCFAAF');
  Ini.Free;
  //======解密password
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

  lytray1.Hint:='数据接收服务'+ExtractFileName(Application.ExeName);

//=============================初始化密码=====================================//
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
        //MessageBox(application.Handle,pchar('感谢您使用智能监控系统，'+chr(13)+'请记住初始化密码：'+'lc'),
        //            '系统提示',MB_OK+MB_ICONinformation);     //WARNING
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

  autorun:=ini.readBool(IniSection,'开机自动运行',false);
  ifRecLog:=ini.readBool(IniSection,'调试日志',false);

  ValueFile:=ini.ReadString(IniSection,'接口文件','');

  GroupName:=trim(ini.ReadString(IniSection,'工作组',''));
  EquipChar:=trim(uppercase(ini.ReadString(IniSection,'仪器字母','')));//读出来是大写就万无一失了
  SpecType:=ini.ReadString(IniSection,'默认样本类型','');
  SpecStatus:=ini.ReadString(IniSection,'默认样本状态','');
  CombinID:=ini.ReadString(IniSection,'组合项目代码','');

  LisFormCaption:=ini.ReadString(IniSection,'检验系统窗体标题','');

  QuaContSpecNoG:=ini.ReadString(IniSection,'高值质控联机号','9999');
  QuaContSpecNo:=ini.ReadString(IniSection,'常值质控联机号','9998');
  QuaContSpecNoD:=ini.ReadString(IniSection,'低值质控联机号','9997');

  SdfDateFormat:=ini.ReadString(IniSection,'结果文件名的日期格式','YYYYMMDD');

  ini.Free;

  OperateLinkFile(application.ExeName,'\'+ChangeFileExt(ExtractFileName(Application.ExeName),'.lnk'),15,autorun);

  if FileExists(ValueFile) then
  BEGIN
    INI_SRC_LIS:=TINIFILE.Create(ValueFile);
    path_result:=INI_SRC_LIS.ReadString('LIS','PATH_RESULT','C:\');
    big_result:=INI_SRC_LIS.ReadString('LIS','big_result',',');
    INI_SRC_LIS.Free;
    Timer1.Enabled:=true;
  END else memo1.Lines.Add('没有找到文件'+ValueFile); 
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
    ss:='服务器'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '数据库'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '集成登录模式'+#2+'CheckListBox'+#2+#2+'0'+#2+#2+#3+
        '用户'+#2+'Edit'+#2+#2+'0'+#2+#2+#3+
        '口令'+#2+'Edit'+#2+#2+'0'+#2+#2+'1';
    if ShowOptionForm('连接数据库','连接数据库',Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
      goto labReadIni else application.Terminate;
  end;
end;

procedure TfrmMain.ToolButton2Click(Sender: TObject);
var
  ss:string;
begin
  if LoadInputPassDll then
  begin
    ss:='接口文件'+#2+'File'+#2+#2+'1'+#2+'注:一般为\Laboman4.0\lis_interface.ini'+#2+#3+
      '结果文件名的日期格式'+#2+'Combobox'+#2+'YYYYMMDD'+#13+'YYYYMD'+#2+'0'+#2+'日期2015年1月20日,YYYYMMDD->20150120,YYYYMD->2015120'+#2+#3+
      '工作组'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '仪器字母'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '检验系统窗体标题'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '默认样本类型'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '默认样本状态'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '组合项目代码'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '开机自动运行'+#2+'CheckListBox'+#2+#2+'1'+#2+#2+#3+
      '调试日志'+#2+'CheckListBox'+#2+#2+'0'+#2+'注:强烈建议在正常运行时关闭'+#2+#3+
      '高值质控联机号'+#2+'Edit'+#2+#2+'2'+#2+#2+#3+
      '常值质控联机号'+#2+'Edit'+#2+#2+'2'+#2+#2+#3+
      '低值质控联机号'+#2+'Edit'+#2+#2+'2'+#2+#2;

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
  showmessage('保存成功!');
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
  ss:='RegisterNum'+#2+'Edit'+#2+#2+'0'+#2+'将该窗体标题栏上的字符串发给开发者,以获取注册码'+#2;
  if bRegister then exit;
  if ShowOptionForm(Pchar('注册:'+GetHDSn('C:\')+'-'+GetHDSn('D:\')+'-'+ChangeFileExt(ExtractFileName(Application.ExeName),'')),Pchar(IniSection),Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
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

  //图形路径
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
  i0:TDateTime;//上次检验时间
  i1:TDateTime;//本次检验时间
  sName:string;//文件名
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
  if ls.Count<=0 then begin ls.Free;exit;end;//如果仪器还没向cdf文件中写完，则等待写完

  //本次检验时间
  i1:=1;
  for i :=0  to ls.Count-1 do
  begin
    lsValue:=StrToList(ls[i],big_result);//将每行导入到字符串列表中

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

  if i1<=i0 then begin ls.Free;exit;end;//该文件已经处理过或已处理过以前做的
  
  ini:=TINIFILE.Create(ChangeFileExt(Application.ExeName,'.ini'));
  ini.WriteDateTime(FormatDateTime('YYYYMMDD',now),s1,i1);
  ini.Free;

  if length(frmMain.memo1.Lines.Text)>=60000 then frmMain.memo1.Lines.Clear;//memo只能接受64K个字符
  frmMain.memo1.Lines.Add(filename);

  //取图形数据
  for i :=0  to ls.Count-1 do
  begin
    lsValue:=StrToList(ls[i],big_result);//将每行导入到字符串列表中

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
    lsValue:=StrToList(ls[i],big_result);//将每行导入到字符串列表中

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
      ifRecLog,true,'常规');
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
        MessageBox(application.Handle,pchar('该程序已在运行中！'),
                    '系统提示',MB_OK+MB_ICONinformation);   
        Halt;
    end;

finalization
    if hnd <> 0 then CloseHandle(hnd);

end.
