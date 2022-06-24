//Comandos SQL Access
//https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/drop-statement-microsoft-access-sql

unit uMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, System.StrUtils,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Data.Win.ADODB, Vcl.Grids,
  Vcl.DBGrids, Vcl.StdCtrls, Vcl.Buttons, Vcl.ComCtrls, uAccessUtil;

type
  TfrmMain = class(TForm)
    ADOConnection: TADOConnection;
    ADOQuery: TADOQuery;
    DataSource: TDataSource;
    cbTabelas: TComboBox;
    cbCampos: TComboBox;
    Label1: TLabel;
    Label2: TLabel;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    tabEstrutura: TTabSheet;
    Memo1: TMemo;
    btnExecuta: TBitBtn;
    DBGrid1: TDBGrid;
    DBGrid2: TDBGrid;
    QryEstrutura: TADOQuery;
    Label3: TLabel;
    cbComando: TComboBox;
    btnGerar: TBitBtn;
    CheckBox1: TCheckBox;
    BitBtn1: TBitBtn;
    dsEstrutura: TDataSource;
    qryUtil: TADOQuery;
    QryEstruturaCODIGO: TIntegerField;
    QryEstruturaTabela: TWideStringField;
    QryEstruturaNomeCampo: TWideStringField;
    QryEstruturaTipoCampo: TWideStringField;
    QryEstruturaTamanho: TIntegerField;
    edtBancoPath: TEdit;
    BitBtn2: TBitBtn;
    OpenDialog1: TOpenDialog;
    SaveDialog1: TSaveDialog;
    procedure btnExecutaClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure cbCamposEnter(Sender: TObject);
    procedure PageControl1Change(Sender: TObject);
    procedure btnGerarClick(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure cbTabelasChange(Sender: TObject);
    procedure BitBtn2Click(Sender: TObject);
  private
    { Private declarations }
    fBaseDados : String;
    function existeTabela(pNomeTabela:string):Boolean;
    Procedure PopularBoxTabelas();
    Procedure PopularBoxCampo();
    Procedure PopularBoxCampoInverso();
    Procedure ConectarBanco();
    procedure ControleConectado(pHabilita: Boolean);
  public
    { Public declarations }
    procedure gerarTabelaEstrutura();
  end;

var
  frmMain: TfrmMain;

implementation

{$R *.dfm}


procedure TfrmMain.FormShow(Sender: TObject);
begin
  PageControl1.ActivePageIndex := 0;
  fBaseDados := LeIni();

  if fBaseDados<>'' then
  begin
    edtBancoPath.Text := fBaseDados;
    ConectarBanco();
    PopularBoxTabelas();
    cbCampos.Items.Clear;

    cbComando.ItemIndex := 0;
    gerarTabelaEstrutura();//o combo precisa está alimentado com as tabelas
  end
  else
  begin
    controleConectado(False);
    Showmessage('Base de dados não foi definida!');
  end;
end;

procedure TfrmMain.btnGerarClick(Sender: TObject);
var
  sComando, sTabela, sCampo: string;
begin
  sComando := trim(cbComando.Items[cbComando.ItemIndex]);
  sTabela  := cbTabelas.Items[cbTabelas.ItemIndex];
  sCampo   := cbCampos.Items[cbCampos.ItemIndex];

  if not (StartsText('Criar Tabela', sComando)) then
     if Trim(cbTabelas.Text)='' then exit;

   Memo1.Lines.Clear;

  if (StartsText('Selecionar', sComando)) then
  begin
    if cbCampos.Text='' then
      Memo1.Text := 'SELECT * FROM '+sTabela+';'
    else
      Memo1.Text := 'SELECT * FROM '+sTabela+' WHERE '+cbCampos.Text+' = '''';';
    exit;
  end;

  if (StartsText('Excluir Dados', sComando)) then
  begin
    if cbCampos.Text='' then
      Memo1.Text := 'DELETE FROM '+sTabela+';'
    else
      Memo1.Text := 'DELETE FROM '+sTabela+' WHERE '+cbCampos.Text+' = '''';';

    exit;
  end;

  if (StartsText('Criar Tabela', sComando)) then
  begin
    Memo1.Lines.Add('CREATE TABLE TabelaNova(CODIGO INTEGER, Empresa INTEGER, Nome VARCHAR(30),'+
                    ' Desconto FLOAT, Valor MONEY, Data DATETIME,'+
                    ' CONSTRAINT [PrimaryKey] PRIMARY KEY (CODIGO, EMPRESA))');
    exit;
  end;

  if (StartsText('Excluir Tabela', sComando)) then
  begin
    Memo1.Text := 'DROP TABLE '+sTabela+';';
    exit;
  end;

  if (StartsText('Adicionar Campo', sComando)) then
  begin
    Memo1.Lines.Add('ALTER TABLE  '+sTabela+' ADD COLUMN NomeColuna1 TEXT(10), NomeColuna2 MEMO');
    exit;
  end;

  if (StartsText('Alterar Campo', sComando)) then
  begin
    if cbCampos.Text<> '' then
       Memo1.Lines.Add('ALTER TABLE '+sTabela+' ALTER COLUMN '+sCampo+' CHAR(20);');
    exit;
  end;

  if (StartsText('Apagar Campo', sComando)) then
  begin
    if cbCampos.Text<> '' then
       Memo1.Lines.Add('ALTER TABLE '+sTabela+' DROP COLUMN '+sCampo+';');
    exit;
  end;


end;


procedure TfrmMain.cbCamposEnter(Sender: TObject);
begin
  if CheckBox1.Checked then
    PopularBoxCampoInverso()
  else
    PopularBoxCampo();
end;

procedure TfrmMain.cbTabelasChange(Sender: TObject);
begin
  PageControl1Change(Self);
  cbCampos.Items.Clear;
end;

procedure TfrmMain.CheckBox1Click(Sender: TObject);
begin
  cbCamposEnter(Sender);
end;

procedure TfrmMain.PageControl1Change(Sender: TObject);
var
  x: integer;
begin
  if cbTabelas.Text='' then
  begin
    ShowMessage('Para ter acesso o design table é necessário selecionar uma tabela na lista!');
    cbTabelas.SetFocus;
    Exit;
  end;

  if PageControl1.ActivePage=tabEstrutura then
  begin
    qryUtil.Close();
    qryUtil.SQL.Text := 'SELECT TOP 1 * FROM '+cbTabelas.Text+';';
    qryUtil.Open;

    QryEstrutura.Close;
    QryEstrutura.SQL.Clear;
    QryEstrutura.SQL.Add('SELECT * FROM ZESTRUTURA01 WHERE Tabela='+QuotedStr(cbTabelas.Text)+';');
    QryEstrutura.Open;

    for x:=0 to qryUtil.FieldCount-1 do
    begin
      if not QryEstrutura.Locate('Tabela;NomeCampo',VarArrayOf([cbTabelas.Text, qryUtil.Fields[x].FieldName]), [loCaseInsensitive]) then
      begin
        QryEstrutura.Append;
        QryEstruturaCODIGO.AsInteger   := x+1;
        QryEstruturaTabela.AsString    := cbTabelas.Text;
        QryEstruturaNomeCampo.AsString := qryUtil.Fields[x].FieldName;
        QryEstruturaTipoCampo.AsString := FieldTypeToStr(qryUtil.Fields[x].DataType);
        QryEstruturaTamanho.AsInteger  := qryUtil.Fields[x].Size;
        QryEstrutura.Post;
      end
      else
      begin
        if (QryEstruturaTipoCampo.AsString <> FieldTypeToStr(qryUtil.Fields[x].DataType))
        or (QryEstruturaTamanho.AsInteger <> qryUtil.Fields[x].Size) then
        begin
          QryEstrutura.Edit;
          QryEstruturaTipoCampo.AsString := FieldTypeToStr(qryUtil.Fields[x].DataType);
          QryEstruturaTamanho.AsInteger  := qryUtil.Fields[x].Size;
          QryEstrutura.Post;
        end;
      end;
    end;
    QryEstrutura.First;
  end;
end;

procedure TfrmMain.BitBtn1Click(Sender: TObject);
begin
  if ADOQuery.Active then
  begin
    if ADOQuery.State IN [dsInsert, dsEdit] then
    begin
      ADOQuery.Edit;
      ADOQuery.Post;
    end;
  end;
end;

procedure TfrmMain.BitBtn2Click(Sender: TObject);
var
  Ligacao : string;
begin
  if trim(edtBancoPath.Text)='' then
    OpenDialog1.InitialDir := ExtractFilePath(ParamStr(0))
  else
    OpenDialog1.InitialDir := ExtractFilePath(edtBancoPath.Text);

  if OpenDialog1.Execute() then
  begin
    edtBancoPath.Text := OpenDialog1.FileName;
    ConectarBanco();

    GravaIni(edtBancoPath.Text);
  end;
end;

procedure TfrmMain.btnExecutaClick(Sender: TObject);
var
  x : Integer;
  sComando : string;
begin
  if trim(Memo1.Text)='' then exit;

  if not isComandoValido(Memo1.Text) then
  begin
    ShowMessage('Comando não é válido!');
    Exit;
  end;

  if not executaComandoCritico(Memo1.Text) then Exit;

  ADOQuery.Close;
  ADOQuery.SQL.Clear;
  ADOQuery.SQL.Add(Memo1.Text);

  sComando := trim(Memo1.Text);
  if (StartsText('SELECT', sComando)) then
    ADOQuery.Open
  else
    ADOQuery.ExecSQL;

  for x:=0 to DBGrid1.FieldCount-1 do
  begin
    if DBGrid1.Fields[x].Size>15 then
      DBGrid1.Fields[x].DisplayWidth := 15;
  end;

  PopularBoxTabelas();
end;

function TfrmMain.existeTabela(pNomeTabela:string):Boolean;
var
  x: integer;
begin
  RESULT := false;
  for x:= 0 to cbTabelas.Items.Count-1 do
  begin
    if cbTabelas.Items[x]=pNomeTabela then
    begin
      RESULT := True;
      Break;
    end;
  end;
end;

procedure TfrmMain.gerarTabelaEstrutura();
begin
  if not existeTabela('ZESTRUTURA01') then
  begin
    QryEstrutura.Close();
    QryEstrutura.SQL.Clear;
    //QryEstrutura.SQL.Add('CREATE TABLE ZESTRUTURA01(CODIGO INTEGER CONSTRAINT CodigoConstraint PRIMARY KEY, Tabela VARCHAR(30), NomeCampo VARCHAR(30), TipoCampo VARCHAR(25), Tamanho INTEGER)');
    QryEstrutura.SQL.Add('CREATE TABLE ZESTRUTURA01(CODIGO INTEGER, Tabela VARCHAR(30), NomeCampo VARCHAR(30), TipoCampo VARCHAR(25), Tamanho INTEGER,CONSTRAINT [PrimaryKey] PRIMARY KEY (CODIGO,Tabela))');

    QryEstrutura.ExecSQL;
    PopularBoxTabelas();
  end;
end;

Procedure TfrmMain.PopularBoxTabelas();
var
  lstTabelas: TStrings;
  x, iindex: integer;
begin
  iindex := cbTabelas.ItemIndex;

  cbTabelas.Items.Clear;
  lstTabelas := TStringList.Create;
  try
    ADOConnection.GetTableNames(lstTabelas);

    for x := 0 to lstTabelas.Count-1 do
      cbTabelas.Items.Add(lstTabelas[x]);

  finally
    lstTabelas.Free;
  end;

  cbTabelas.ItemIndex := iindex;
end;

Procedure TfrmMain.PopularBoxCampo();
var
  lstCampos: TStrings;
  x: integer;
begin
  cbCampos.Items.Clear;
  lstCampos := TStringList.Create;
  try
    ADOConnection.GetFieldNames(cbTabelas.Items[cbTabelas.ItemIndex], lstCampos);

    for x := 0 to lstCampos.Count-1 do
    begin
      cbCampos.Items.Add(lstCampos[x]);
    end;

  finally
    lstCampos.Free;
  end;
end;

Procedure TfrmMain.PopularBoxCampoInverso();
var
  lstCampos: TStrings;
  x: integer;
begin
  cbCampos.Items.Clear;
  lstCampos := TStringList.Create;
  try
    ADOConnection.GetFieldNames(cbTabelas.Items[cbTabelas.ItemIndex], lstCampos);

    x := lstCampos.Count;
    while x>0 do
    begin
       cbCampos.Items.Add(lstCampos[x-1]);
       x := x-1;
    end;

  finally
    lstCampos.Free;
  end;
end;

Procedure TfrmMain.ConectarBanco();
var
  Ligacao: string;
begin
    Ligacao :=
    'Provider=Microsoft.Jet.OLEDB.4.0;'+
    'User ID=Admin;'+
    'Data Source='+edtBancoPath.Text+';'+
    'Mode=Share Deny None;Persist Security Info=False;'+
    'Jet OLEDB:System database="";'+
    'Jet OLEDB:Registry Path="";'+
    'Jet OLEDB:Database Password="";'+
    'Jet OLEDB:Engine Type=5;'+
    'Jet OLEDB:Database Locking Mode=1;'+
    'Jet OLEDB:Global Partial Bulk Ops=2;'+
    'Jet OLEDB:Global Bulk Transactions=1;'+
    'Jet OLEDB:New Database Password="";'+
    'Jet OLEDB:Create System Database=False;'+
    'Jet OLEDB:Encrypt Database=False;'+
    'Jet OLEDB:Don''t Copy Locale on Compact=False;'+
    'Jet OLEDB:Compact Without Replica Repair=False;'+
    'Jet OLEDB:SFP=False';

    ADOConnection.Connected        := false;
    ADOConnection.LoginPrompt      := false;
    ADOConnection.ConnectionString := Ligacao;
    ADOConnection.Connected        := true;

    controleConectado(true);
end;

procedure TfrmMain.ControleConectado(pHabilita: Boolean);
begin
  btnExecuta.Enabled := pHabilita;
  cbTabelas.Enabled  := pHabilita;
  cbCampos.Enabled   := pHabilita;
end;


end.

//[AUTO]                  COUNTER
//[Byte]                  BYTE
//[Integer]               SMALLINT
//[Long]                  INTEGER
//[Single]                REAL
//[Double]                FLOAT
//[Decimal]               DECIMAL(18,5)
//[Currency]              MONEY
//[ShortText]             VARCHAR
//[LongText]              MEMO
//[DateTime]              DATETIME
//[YesNo]                 BIT
//[OleObject]             IMAGE
//[Required]              INTEGER NOT NULL

