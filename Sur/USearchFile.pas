unit USearchFile;

interface

uses
  SysUtils{TSearchRec},Forms{application};

type 
  {TFindCallBackΪ�ص�����,FindFile�����ҵ�һ��ƥ����ļ�֮��ͻ�����������.
  filename:�ҵ����ļ���(����·��),���ڻص������п��Ը����ļ������в���.
  info:�ҵ����ļ��ļ�¼��Ϣ,��һ��TSearchRec�ṹ.
  bQuit:�����Ƿ���ֹ�ļ��Ĳ���}
  TFindCallBack=procedure(const filename:string;const info:TSearchRec;var bQuit:boolean);

procedure FindFile(var Quit:boolean;const Path:String;const filename:string='*.*'; 
                  proc:TFindCallBack=nil;const bSub:boolean=true;const bMsg:boolean=true); 

implementation


procedure FindFile(var Quit:boolean;const Path:String;const filename:string='*.*';
                  proc:TFindCallBack=nil;const bSub:boolean=true;const bMsg:boolean=true); 
//Quit:�����Ƿ��˳����ң�Ӧ�ó�ʼ��Ϊfalse�� 
//Path:����·���� 
//filename:�ļ��������԰���Windows��֧�ֵ��κ�ͨ����ĸ�ʽ��Ĭ�����е��ļ� 
//proc:�ص�������Ĭ��Ϊ�� 
//bSub:�����Ƿ������Ŀ¼��Ĭ��Ϊ������Ŀ¼ 
//bMsg:�����Ƿ��ڲ����ļ���ʱ������������Ϣ��Ĭ��Ϊ������������Ϣ 
var 
  fpath: String;
  info: TsearchRec; 

  procedure ProcessAFile; 
  begin 
    if (info.Name<>'.')and(info.Name<>'..')and((info.Attr and faDirectory)<>faDirectory) then 
    begin 
      if assigned(proc) then proc(fpath+info.FindData.cFileName,info,quit);
    end; 
  end; 

  procedure ProcessADirectory; 
  begin 
    if (info.Name<>'.')and(info.Name<>'..')and((info.attr and fadirectory)=fadirectory) then
      findfile(quit,fpath+info.Name,filename,proc,bsub,bmsg); 
  end; 

begin 
  if path[length(path)]<>'\' then fpath:=path+'\' else fpath:=path;
  try 
    if 0=findfirst(fpath+filename,faanyfile and (not fadirectory),info) then 
    begin 
     ProcessAFile; 
     while 0=findnext(info) do 
     begin 
       ProcessAFile; 
       if bmsg then application.ProcessMessages; 
       if quit then 
       begin 
         findclose(info); 
         exit; 
       end; 
     end; 
    end; 
  finally 
    findclose(info); 
  end; 
  try 
    if bsub and (0=findfirst(fpath+'*',faanyfile,info)) then
    begin
      ProcessADirectory;
      while findnext(info)=0 do ProcessADirectory;
    end;
  finally
    findclose(info); 
  end;
end;

{����ʾ��
procedure AFindCallBack(const filename:string;const info:tsearchrec;var quit:boolean);
begin 
 form1.listbox1.Items.Add(filename);
end;

procedure TForm1.Button10Click(Sender: TObject);
begin
  listbox1.Clear;
  qqq:=false;
  findfile(qqq,edit1.text,'*.*',AFindCallBack,true,true);
end;

procedure TForm1.Button11Click(Sender: TObject);
begin
  qqq:=true;
end;
//}

end.
