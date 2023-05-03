unit uexcellink;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, Grids,
  ComObj, ShellAPI, Windows;

type

  { TForm1 }

  TForm1 = class(TForm)
    Button1: TButton;
    Button10: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    Button5: TButton;
    Button6: TButton;
    Button7: TButton;
    Button8: TButton;
    Button9: TButton;
    CheckBox1: TCheckBox;
    Edit1: TEdit;
    Edit2: TEdit;
    Edit3: TEdit;
    Edit4: TEdit;
    Edit5: TEdit;
    Edit6: TEdit;
    Edit7: TEdit;
    Edit8: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Memo1: TMemo;
    OpenDialog1: TOpenDialog;
    SG: TStringGrid;
    procedure Button10Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure Button9Click(Sender: TObject);
    procedure fillsgindexes;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure OpenCSVFile(FileName: String; separator: char);
    procedure OpenExcelFile;
    procedure SGClick(Sender: TObject);
  private

  public

  end;

var
  Form1: TForm1;
  XLApp: OLEVariant;
  x,y: byte;
  path: variant;
  ldata:TStringList;

implementation

procedure tform1.fillsgindexes;
var i,e:integer;
begin
  e:=SG.RowCount;
  for i:=1 to e-1 do
  begin
    SG.Cells[0,i]:=IntToStr(i);
  end;
end;

procedure TForm1.Button8Click(Sender: TObject);
var nr,nc:integer;
    r,c:byte;
begin
 XLApp := CreateOleObject('Excel.Application'); // requires comobj in uses
 try
   XLApp.Visible := True;         // Show Excel
   XLApp.DisplayAlerts := True;
   XLApp.Workbooks.Add;     // Open the Workbook
   nr:=SG.RowCount-1;
   nc:=SG.ColCount-1;
   Memo1.Clear;
   for c := 1 to nc do
   begin
     for r := 1 to nr do
      begin
       Memo1.Lines.Add('Writing: r='+inttostr(r)+'; c='+inttostr(c)+'; val='+sg.Cells[c,r]);
       XLApp.Cells[r,c].Value:=SG.Cells[c,r];  // fill spreadsheet with values
       Memo1.Lines.Add('Written: r='+inttostr(r)+'; c='+inttostr(c)+'; val='+sg.Cells[c,r]);
      end;
   end;
  except
    ShowMessage('Something went wrong');
  end;
end;

procedure TForm1.Button10Click(Sender: TObject);
begin
  SG.Cells[SG.Selection.Left,SG.Selection.Top]:=Edit4.Text;
end;

procedure TForm1.Button9Click(Sender: TObject);
begin
  SG.ColCount:=strtoint(edit7.text)+1;
  SG.RowCount:=strtoint(edit8.text)+1;
end;

procedure TForm1.OpenCSVFile(FileName: String; separator: char);
var i,j,n,sgr:integer;
    s,k:string;
begin
  ldata:=TStringList.Create();
  ldata.LoadFromFile(FileName);
  SG.Clear;
  n:=ldata.Count;
  SG.RowCount:=n+1;
  if n>0 then
  for i:=0 to (n-1) do
  begin
    s:=ldata[i];
    k:='';
    sgr:=1;
    for j:=1 to length(s) do
    begin
      if s[j]<>separator then k:=k+s[j];
      if (s[j]=separator) or (j=length(s)) then
      begin
        if SG.ColCount<(sgr+1) then SG.ColCount:=sgr+1;
        SG.Cells[sgr,i+1]:=k;
        inc(sgr);
        k:='';
      end;
    end;
  end;

end;

procedure TForm1.OpenExcelFile;
var n,l,counter:integer;
begin
 XLApp := CreateOleObject('Excel.Application'); // requires comobj in uses
 try
   XLApp.Visible := False;         // Hide Excel
   XLApp.DisplayAlerts := False;
   path := edit1.Text;
   XLApp.Workbooks.Open(Path);     // Open the Workbook
   SG.Clear;
   n:=XLApp.Sheets[1].UsedRange.Rows.Count;
   l:=XLApp.Sheets[1].UsedRange.Columns.Count;
   SG.RowCount:=n+1;
   SG.ColCount:=l+1;
   counter:=0;
   for x :=1 to (l) do
   begin
     for y := 1 to (n) do
      begin
       SG.Cells[0,y] := inttostr(counter);
       SG.Cells[x,y] := XLApp.Cells[y,x].Value;  // fill stringgrid with values
       inc(counter);
      end;
   end;
   SG.Cells[0,0]:='id';
   SG.Cells[1,0]:='X';
   SG.Cells[2,0]:='Y';
  finally
   XLApp.Quit;
   XLAPP := Unassigned;
  end;
end;

procedure TForm1.SGClick(Sender: TObject);
begin
  try
  if GetKeyState(VK_SHIFT)<0 then
  begin
    edit6.Text:=inttostr(SG.Selection.Bottom);
  end
  else
  begin
    Edit2.Text:=inttostr(SG.Selection.Left);
    Edit3.Text:=inttostr(SG.Selection.Top);
    Edit4.Text:=SG.Cells[SG.Selection.Left,SG.Selection.Top];
  end;
  except
  end;
end;

procedure sshellexecurl(url:string);
var turl:string;
begin
    turl := StringReplace(URL, '"', '%22', [rfReplaceAll]);
    ShellExecute(0, 'open', PChar(turl), nil, nil, 1);
end;

procedure TForm1.Button4Click(Sender: TObject);
var url:string;
    i,c,r1,r2:integer;
begin
  if edit3.Text=edit6.text then
  begin
    URL := edit4.text;
    sshellexecurl(url);
  end
  else
  begin
     c:=strtoint(edit2.Text);
     r1:=strtoint(edit3.text);
     r2:=strtoint(edit6.text);
     if r2>r1 then
     begin
       for i:=r1 to r2 do
       begin
         url:=SG.Cells[c,i];
         sshellexecurl(url);
       end;
     end;
  end;
end;

procedure TForm1.Button5Click(Sender: TObject);
begin
  edit4.text:=SG.Cells[strtoint(edit2.text),strtoint(edit3.text)];
end;

procedure TForm1.Button6Click(Sender: TObject);
begin
  OpenCSVFile(edit1.text,edit5.text[1]);
  fillsgindexes;
end;

procedure TForm1.Button7Click(Sender: TObject);
var d,i,c,r1,r2:integer;
begin
  c:=strtoint(edit2.Text);
  r1:=strtoint(edit3.text);
  r2:=strtoint(edit6.text);
  d:=r2-r1;
  r1:=r1+d+1;
  r2:=r2+d+1;
  edit3.text:=inttostr(r1);
  edit6.text:=inttostr(r2);
  SG.Selection.SetLocation(c,r1);
  SG.Selection.Offset(0,d+1);

end;

procedure TForm1.FormResize(Sender: TObject);
begin
  SG.Width:=Form1.ClientWidth-(Button4.Left+Button4.Width+10)-Memo1.Width;
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  OpenExcelFile;
  fillsgindexes;
end;

procedure TForm1.Button3Click(Sender: TObject);
var a:integer;
begin
  a:=strtoint(edit3.text);
  a:=a+1;
  edit3.text:=inttostr(a);
  if CheckBox1.Checked then Button5.Click;
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
  if OpenDialog1.Execute then edit1.text:=opendialog1.filename;
end;

{$R *.lfm}

end.

