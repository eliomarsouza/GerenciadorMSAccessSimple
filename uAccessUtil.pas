unit uAccessUtil;

interface

uses System.SysUtils, Vcl.Forms, Winapi.Windows, System.StrUtils, Data.DB, Dialogs;


  procedure GravaIni(aTexto: string);
  function LeIni():String;
  function FieldTypeToStr(ftCampo: TFieldType) : string;
  function isComandoValido(psComando: string): Boolean;
  function executaComandoCritico(psComando: string): Boolean;

  const ArquivoIni='ConfigAccess.ini';


implementation

uses IniFiles;


procedure GravaIni(aTexto: string);
var
  ArqIni: TIniFile;
begin
  ArqIni := TIniFile.Create(ExtractFilePath(ParamStr(0))+ArquivoIni);
  try

    ArqIni.WriteString('Dados', 'BANCO', aTexto);

  finally

    ArqIni.Free;

  end;

end;

function LeIni():String;
var
  ArqIni: TIniFile;
begin
  ArqIni := TIniFile.Create(ExtractFilePath(ParamStr(0))+ArquivoIni);

  try
    Result:=ArqIni.ReadString('Dados', 'BANCO', '');
  finally
    ArqIni.Free;
  end;
end;


function FieldTypeToStr(ftCampo: TFieldType) : string;
begin
  case ftCampo of
    ftAutoInc:
      Result := 'AutoInc';
    ftString,ftWideString:
       Result := 'Varchar';
    ftInteger,ftShortint,ftSmallint,ftLargeint:
       Result := 'Int';
    ftCurrency, ftBCD, ftFMTBcd:
       Result := 'Money';
    ftDateTime:
       Result := 'Datetime';
    ftDate:
       Result := 'Date';
    ftTimeStamp:
       Result := 'TimeStamp';
    ftFloat:
       Result := 'Float';
    ftBoolean:
       Result := 'Bit';
    ftMemo,ftWideMemo,ftFmtMemo:
       Result := 'Memo';
    ftBlob:
       Result := 'Blob';
    ftVariant:
      Result := 'Variant';
    else
      Result := 'Não achou tipo';
  end;

  {ftUnknown, ftBytes, ftVarBytes, ftGraphic,
  ftParadoxOle, ftDBaseOle, ftTypedBinary, ftCursor, ftFixedChar,
  ftADT, ftArray, ftReference, ftDataSet, ftOraBlob, ftOraClob,
  ftVariant, ftInterface, ftIDispatch, ftGuid,
  ftFixedWideChar, ftOraTimeStamp, ftOraInterval,
  ftLongWord, ftByte, ftExtended, ftConnection, ftParams, ftStream,
  ftTimeStampOffset, ftObject, ftSingle}
end;

function isComandoValido(psComando: string): Boolean;
begin
  psComando := UpperCase(trim(psComando));
  result    := StartsText('SELECT',psComando)
            or StartsText('INSERT',psComando)
            or StartsText('DELETE',psComando)
            or StartsText('ALTER TABLE',psComando)
            or StartsText('DROP TABLE',psComando)
            or StartsText('CREATE',psComando)
            ;
end;

function executaComandoCritico(psComando: string): Boolean;
var
  bCritico: Boolean;
begin
  result := true;

  psComando := UpperCase(trim(psComando));

  bCritico  := StartsText('DROP',psComando)
            or StartsText('DELETE',psComando)
            or StartsText('ALTER',psComando)
            ;

  if bCritico then
    result := Application.MessageBox('Perigo comando crítico, deseja gera-lo mesmo assim?','',MB_ICONWARNING+MB_YESNO)=IDYES;
end;

end.
