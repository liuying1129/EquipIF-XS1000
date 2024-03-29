unit UfrmMain;

interface

uses
  Windows, Messages, SysUtils, Classes, Controls, Forms,
  Menus, StdCtrls, Buttons, ADODB,
  ComCtrls, ToolWin, ExtCtrls,
  inifiles,Dialogs,
  StrUtils, DB,ComObj,Variants, CoolTrayIcon;

type
  TfrmMain = class(TForm)
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    ADOConnection1: TADOConnection;
    CoolBar1: TCoolBar;
    ToolBar1: TToolBar;
    ToolButton7: TToolButton;
    ToolButton8: TToolButton;
    ToolButton2: TToolButton;
    Memo1: TMemo;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    Button1: TButton;
    ToolButton5: TToolButton;
    ToolButton9: TToolButton;
    OpenDialog1: TOpenDialog;
    Timer1: TTimer;
    LYTray1: TCoolTrayIcon;
    procedure N3Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    //增加病人信息表中记录,返回该记录的唯一编号作为检验结果表的外键
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure N1Click(Sender: TObject);
    procedure ToolButton7Click(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure ToolButton5Click(Sender: TObject);
  private
    { Private declarations }
    procedure UpdateConfig;{配置文件生效}
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
  //SdfDateFormat:string;//结果文件名的日期格式
  ifRecLog:boolean;//是否记录调试日志
  ExcludeLJBS:STRING;//排除联机标识
  DataFileExtension:string;//数据文件后缀
  EquipUnid:integer;//设备唯一编号

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
  //Persist Security Info,表示ADO在数据库连接成功后是否保存密码信息
  //ADO缺省为True,ADO.net缺省为False
  //程序中会传ADOConnection信息给TADOLYQuery,故设置为True
  result := result + 'Persist Security Info=True;';
  if ifIntegrated then
    result := result + 'Integrated Security=SSPI;';
end;

procedure TfrmMain.FormCreate(Sender: TObject);
begin
  ConnectString:=GetConnectString;
  
  UpdateConfig;
  if ifRegister then bRegister:=true else bRegister:=false;  

  Caption:='数据接收服务'+ExtractFileName(Application.ExeName);
  lytray1.Hint:='数据接收服务'+ExtractFileName(Application.ExeName);
end;

procedure TfrmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  action:=caNone;
  LYTray1.HideMainForm;
end;

procedure TfrmMain.N3Click(Sender: TObject);
begin
  if (MessageDlg('退出后将不再接收设备数据,确定退出吗？', mtWarning, [mbYes, mbNo], 0) <> mrYes) then exit;
  application.Terminate;
end;

procedure TfrmMain.N1Click(Sender: TObject);
begin
  LYTray1.ShowMainForm;
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

  DataFileExtension:=trim(ini.ReadString(IniSection,'数据文件后缀','cdf'));
  if DataFileExtension='' then DataFileExtension:='cdf';
  GroupName:=trim(ini.ReadString(IniSection,'工作组',''));
  EquipChar:=trim(uppercase(ini.ReadString(IniSection,'仪器字母','')));//读出来是大写就万无一失了
  SpecType:=ini.ReadString(IniSection,'默认样本类型','');
  SpecStatus:=ini.ReadString(IniSection,'默认样本状态','');
  CombinID:=ini.ReadString(IniSection,'组合项目代码','');

  LisFormCaption:=ini.ReadString(IniSection,'检验系统窗体标题','');
  ExcludeLJBS:=ini.ReadString(IniSection,'排除联机标识','');
  EquipUnid:=ini.ReadInteger(IniSection,'设备唯一编号',-1);

  QuaContSpecNoG:=ini.ReadString(IniSection,'高值质控联机号','9999');
  QuaContSpecNo:=ini.ReadString(IniSection,'常值质控联机号','9998');
  QuaContSpecNoD:=ini.ReadString(IniSection,'低值质控联机号','9997');

  //SdfDateFormat:=ini.ReadString(IniSection,'结果文件名的日期格式','YYYYMMDD');

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
    ss:='接口文件'+#2+'File'+#2+#2+'1'+#2+'注:一般为\Laboman4.0\lis_interface.ini'+#2+#3+
      //'结果文件名的日期格式'+#2+'Combobox'+#2+'YYYYMMDD'+#13+'YYYYMD'+#2+'0'+#2+'日期2015年1月20日,YYYYMMDD->20150120,YYYYMD->2015120'+#2+#3+
      '数据文件后缀'+#2+'Edit'+#2+#2+'1'+#2+'一般的,填写cdf或sdf,默认值cdf'+#2+#3+
      '工作组'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '仪器字母'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '检验系统窗体标题'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '默认样本类型'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '默认样本状态'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '组合项目代码'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '开机自动运行'+#2+'CheckListBox'+#2+#2+'1'+#2+#2+#3+
      '调试日志'+#2+'CheckListBox'+#2+#2+'0'+#2+'注:强烈建议在正常运行时关闭'+#2+#3+
      '排除联机标识'+#2+'Edit'+#2+#2+'1'+#2+'多个联机标识用逗号分隔'+#2+#3+
      '设备唯一编号'+#2+'Edit'+#2+#2+'1'+#2+#2+#3+
      '高值质控联机号'+#2+'Edit'+#2+#2+'2'+#2+#2+#3+
      '常值质控联机号'+#2+'Edit'+#2+#2+'2'+#2+#2+#3+
      '低值质控联机号'+#2+'Edit'+#2+#2+'2'+#2+#2;

  if ShowOptionForm('',Pchar(IniSection),Pchar(ss),Pchar(ChangeFileExt(Application.ExeName,'.ini'))) then
	  UpdateConfig;
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
  ls,lsValue{,sList}:tstrings;
  i:integer;

  SpecNo:string;
  FInts:OleVariant;
  ReceiveItemInfo:OleVariant;

  //ini:Tinifile;

  //s1:string;
  //i0:TDateTime;//上次检验时间
  //i1:TDateTime;//本次检验时间
  //sName:string;//文件名
  //fs:TFormatSettings;
  //s2:string;

  s3:string;
  ls3:tstrings;
  CheckDate:string;
begin
  //sName:=ExtractFileName(filename);
  
  //sList:=TStringList.Create;
  //ExtractStrings(['_'],[],PChar(sName),sList);
  //if sList.Count<3 then begin sList.Free;exit;end;
  //s1:=sList[0]+'_'+sList[1];

  //本次检验时间
  //i1:=1;
  //s2:=FormatDateTime('YYYY-MM-DD',now)+' '+copy(sList[2],1,2)+':'+copy(sList[2],3,2)+':'+copy(sList[2],5,2);
  //fs.DateSeparator:='-';
  //fs.TimeSeparator:=':';
  //fs.ShortDateFormat:='YYYY-MM-DD hh:nn:ss';
  //i1:=StrtoDateTimeDef(s2,i1,fs);
  //==========
  
  //sList.Free;
    
  //ini:=TINIFILE.Create(ChangeFileExt(Application.ExeName,'.ini'));
  //i0:=ini.ReadDateTime(FormatDateTime('YYYYMMDD',now),s1,0);
  //ini.Free;

  ls:=Tstringlist.Create;
  ls.LoadFromFile(filename);
  if ls.Count<=0 then begin ls.Free;exit;end;//如果仪器还没向cdf文件中写完，则等待写完

  //本次检验时间
  {i1:=1;
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
  end;//}
  //==========

  //if i1<=i0 then begin ls.Free;exit;end;//该文件已经处理过或已处理过以前做的
  
  //ini:=TINIFILE.Create(ChangeFileExt(Application.ExeName,'.ini'));
  //ini.WriteDateTime(FormatDateTime('YYYYMMDD',now),s1,i1);
  //ini.Free;

  if length(frmMain.memo1.Lines.Text)>=60000 then frmMain.memo1.Lines.Clear;//memo只能接受64K个字符
  frmMain.memo1.Lines.Add(filename);

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

    if(lsValue[0]='00')and(lsValue.Count>=20) then
      CheckDate:=StringReplace(lsValue[19],'/','-',[rfReplaceAll, rfIgnoreCase]);

    if lsValue[0]='1' then
    begin
      s3:=StringReplace(ExcludeLJBS,'，',',',[rfReplaceAll, rfIgnoreCase]);
      ls3:=TStringList.Create;
      ExtractStrings([','],[],PChar(s3),ls3);
      if ls3.IndexOf(lsValue[1])<0 then
        ReceiveItemInfo[i]:=VarArrayof([lsValue[1],lsValue[3],'',''])
      else ReceiveItemInfo[i]:=VarArrayof(['','','','']);
      ls3.Free;
    end
    else if lsValue[0]='3' then ReceiveItemInfo[i]:=VarArrayof([lsValue[2],'','',lsValue[3]])
    else ReceiveItemInfo[i]:=VarArrayof(['','','','']);

    lsValue.Free;
  end;

  ls.Free;

  if bRegister then
  begin
    FInts :=CreateOleObject('Data2LisSvr.Data2Lis');
    FInts.fData2Lis(ReceiveItemInfo,(SpecNo),CheckDate,
      (GroupName),(SpecType),(SpecStatus),(EquipChar),
      (CombinID),'',(LisFormCaption),(ConnectString),
      (QuaContSpecNoG),(QuaContSpecNo),(QuaContSpecNoD),'',
      ifRecLog,true,'常规',
      '',
      EquipUnid,
      '','','','',
      -1,-1,-1,-1,
      -1,-1,-1,-1,
      false,false,false,false);
    if not VarIsEmpty(FInts) then FInts:= unAssigned;
  end;

  Try
    FileSetAttr(filename,0);//修改文件属性为普通属性,不然可能无法删除
    DeleteFile(filename);//删除文件
  except
    on E:Exception do
    begin
      frmMain.memo1.Lines.Add('设置文件属性或删除文件失败:'+E.Message+'【'+filename+'】');
    end;
  end;
end;

procedure TfrmMain.Timer1Timer(Sender: TObject);
var
  qqq:boolean;

  //lsSection    :TStrings;
  //ini          :Tinifile;
  //i            :integer;
begin
  (Sender as TTimer).Enabled:=false;

  //lsSection:=Tstringlist.Create;
  //ini:=TINIFILE.Create(ChangeFileExt(Application.ExeName,'.ini'));
  //ini.ReadSections(lsSection);
  //for i :=0  to lsSection.Count-1 do
  //begin
  //  if (leftstr(lsSection[i],2)='20')and(lsSection[i]<>FormatDateTime('YYYYMMDD',now)) then ini.EraseSection(lsSection[i]);
  //end;
  //ini.Free;
  //lsSection.Free;

  qqq:=false;
  //findfile(qqq,PATH_RESULT,FormatDateTime(SdfDateFormat,now)+'_*.cdf',AFindCallBack,false,true);
  findfile(qqq,PATH_RESULT,'*_*.'+DataFileExtension,AFindCallBack,false,true);

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
