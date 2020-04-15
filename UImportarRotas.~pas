unit UImportarRotas;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, OleCtnrs, ComObj, Grids, DBXpress, FMTBcd, DB, SqlExpr,
  ComCtrls, Buttons;

type
  TFrmPrincipal = class(TForm)
    btnImportar: TButton;
    edtcaminho: TEdit;
    lblSituacao: TLabel;
    Label2: TLabel;
    connection1: TSQLConnection;
    qryConsulta: TSQLQuery;
    XStringGrid: TStringGrid;
    ProgressBar1: TProgressBar;
    OpenDialog1: TOpenDialog;
    btncarregar: TBitBtn;
    mmLog: TMemo;
    btnCancelar: TButton;
    lblNomeTransportadora: TLabel;
    GroupBox1: TGroupBox;
    lbl1: TLabel;
    combotipoImportacao: TComboBox;
    procedure btnImportarClick(Sender: TObject);
    procedure btncarregarClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnCancelarClick(Sender: TObject);
    procedure edtCodtransportadoraChange(Sender: TObject);
  private
    { Private declarations }
    function Ret_Numero(Key: Char; Texto: string; EhDecimal: Boolean = False): Char;
  public
    { Public declarations }
    procedure cancelarProcessamento();
    function nomeEmpresa(): string;
    function getTipoImportacao(): string;
    function GetNomeComputador: string;
    function validaTransportadora(codigoTransportadora: string; var codiEmpresa: string): Boolean;
  end;

var
  FrmPrincipal: TFrmPrincipal;
  xFileXLS: string;
  cancelar: boolean;

implementation

{$R *.dfm}

procedure TFrmPrincipal.btnImportarClick(Sender: TObject);
const
  xlCellTypeLastCell = $0000000B;
var
  XLSAplicacao, AbaXLS: OLEVariant;
  RangeMatrix: Variant;
  x, y, k, r: Integer;
  empresaValida: Boolean;
  empresa, cnpj, cd_transportadora, cep_inicial, cep_final, uf, cidade, rota, sigla, nomeArquivo, procesado, tipo_importacao, insert, nomeComputador: string;
begin
  if MessageDlg('Deseja realmente importar este arquivo?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    procesado := quotedstr('N');
    nomeComputador := GetNomeComputador;
    tipo_importacao := getTipoImportacao;
    nomeArquivo := edtcaminho.Text;
    cancelar := false;
    btncarregar.Enabled := false;
    btnCancelar.Enabled := true;
    xFileXLS := edtcaminho.Text;
    XLSAplicacao := CreateOleObject('Excel.Application');
    //XLSAplicacao := CreateOleObject('com.sun.star.ServiceManager');
   // Esconde Excel
    XLSAplicacao.Visible := False;
      // Abre o Workbook
    XLSAplicacao.Workbooks.Open(xFileXLS);

      {Selecione aqui a aba que você deseja abrir primeiro - 1,2,3,4....}
    XLSAplicacao.WorkSheets[1].Activate;
      {Selecione aqui a aba que você deseja ativar - começando sempre no 1 (1,2,3,4) }
    AbaXLS := XLSAplicacao.Workbooks[ExtractFileName(xFileXLS)].WorkSheets[1];

    AbaXLS.Cells.SpecialCells(xlCellTypeLastCell, EmptyParam).Activate;
      // Pegar o número da última linha
    x := XLSAplicacao.ActiveCell.Row;
      // Pegar o número da última coluna
    y := XLSAplicacao.ActiveCell.Column;
      // Seta xStringGrid linha e coluna
    XStringGrid.RowCount := x;
    XStringGrid.ColCount := y;
      // Associaca a variant WorkSheet com a variant do Delphi
    RangeMatrix := XLSAplicacao.Range['A1', XLSAplicacao.Cells.Item[x, y]].Value;
      // Cria o loop para listar os registros no TStringGrid
//      cnpj :=Vartostr(RangeMatrix[2, 1]);

    ProgressBar1.Min := 0;
    ProgressBar1.Max := x;

    for r := 2 to x do
    begin
      try
        for k := 1 to y do
        begin
          try
            case k of
              1:
                cd_transportadora := RangeMatrix[r, k];
              2:
                cep_inicial := RangeMatrix[r, k];
              3:
                cep_final := RangeMatrix[r, k];
              4:
                uf := RangeMatrix[r, k];
              5:
                cidade := RangeMatrix[r, k];
              6:
                rota := RangeMatrix[r, k];
              7:
                sigla := RangeMatrix[r, k]
            end;
          except
            cancelarProcessamento;
            mmLog.Enabled := true;
            mmLog.Lines.Add('Erro linha: ' + inttostr(r));
            mmLog.Lines.Add(insert);
          end;
          Inc(y, 1);
        end;
        try
          if validaTransportadora(cd_transportadora, empresa) then
          begin
            ProgressBar1.Position := r;
            lblSituacao.Caption := 'Importando ' + IntToStr(r) + ' de ' + inttostr(x);
            Application.ProcessMessages;
            insert := 'INSERT INTO  T_TRANSPORTADORA_ROTA_IMPORT(CD_EMPRESA  ,CD_TRANSPORTADORA,CD_CEP_INICIO ,CD_CEP_FINAL ,CD_UF' + ' ,DS_MUNICIPIO ,CD_ROTA ,DS_ROTA , ';
            insert := insert + 'DS_SIGLA_ROTEIRIZADOR ,PROCESSADO,TIPO_IMPORTACAO,NOME_ARQUIVO,NOMECOMPUTADOR)values(';

            insert := insert + QuotedStr(empresa) + ',' + quotedstr(cd_transportadora) + ',' + quotedstr(cep_inicial) + ',' + quotedstr(cep_final) + ',' + quotedstr(uf) + ',' + quotedstr(cidade) + ',' + quotedstr(rota) + ',' + quotedstr(cidade) + ',' + quotedstr(sigla) ;
            insert := insert +','+procesado+','+QuotedStr(tipo_importacao)+','+QuotedStr(nomeArquivo)+','+QuotedStr( nomeComputador)+')';
            qryConsulta.Close;
            qryConsulta.SQL.Text := insert;
            qryConsulta.ExecSQL();
          end
          else
          begin
            mmLog.Enabled := true;
            mmLog.Lines.Add('Erro linha: ' + inttostr(r));
            mmLog.Lines.Add(insert);
            Showmessage('Importação Contém erros Entre em contato com a TI');
          end;
          if cancelar then
          begin
            cancelarProcessamento;
            exit;
          end;

        except

          mmLog.Enabled := true;
          mmLog.Lines.Add('Erro linha: ' + inttostr(r));
          mmLog.Lines.Add(insert);
          Showmessage('Importação Contém erros Entre em contato com a TI');

        end;
      except

      end;

    end;
  end;

  lblSituacao.Caption := 'Importação Terminada com sucesso';
  btncancelar.Enabled := false;
  btncarregar.Enabled := true;
end;

procedure TFrmPrincipal.btncarregarClick(Sender: TObject);
begin
  if OpenDialog1.Execute then
  begin
    btnCancelar.Enabled := false;
    edtcaminho.text := OpenDialog1.FileName;
    btnImportar.Enabled := true;
  end;
end;

procedure TFrmPrincipal.FormCreate(Sender: TObject);
begin
  btnImportar.Enabled := false;
  btncarregar.Enabled := true;
  btnCancelar.Enabled := false;
  edtcaminho.Text := '';
end;

procedure TFrmPrincipal.cancelarProcessamento;
begin
  qryConsulta.Close;
  qryConsulta.SQL.Text := 'delete from T_TRANSPORTADORA_ROTA_IMPORT';
  if qryConsulta.ExecSQL() > 0 then
  begin
    ShowMessage('Cancelado com sucesso!!!');
    btnCancelar.Enabled := false;
    btnImportar.Enabled := true;
    btncarregar.Enabled := true;
    ProgressBar1.Position := 0;
    lblSituacao.caption := '';
  end
  else
    ShowMessage('Não foi possivél cancelar contate o TI !!!');
end;

procedure TFrmPrincipal.btnCancelarClick(Sender: TObject);
begin
  cancelar := true;
end;

function TFrmPrincipal.Ret_Numero(Key: Char; Texto: string; EhDecimal: Boolean = False): Char;
begin
  if not EhDecimal then
  begin
      { Chr(8) = Back Space }
    if not (Key in ['0'..'9', Chr(8)]) then
      Key := #0
  end
  else
  begin
    if Key = #46 then
      Key := DecimalSeparator;
    if not (Key in ['0'..'9', Chr(8), DecimalSeparator]) then
      Key := #0
    else if (Key = DecimalSeparator) and (Pos(Key, Texto) > 0) then
      Key := #0;
  end;
  Result := Key;

end;

function TFrmPrincipal.nomeEmpresa: string;
begin

end;

procedure TFrmPrincipal.edtCodtransportadoraChange(Sender: TObject);
begin
  nomeEmpresa;
end;

function TFrmPrincipal.getTipoImportacao: string;
begin
  if combotipoImportacao.ItemIndex = 0 then
    Result := 'A'
  else
    Result := 'N';
end;

function TFrmPrincipal.GetNomeComputador: string;
var
  lpBuffer: PChar;
  nSize: DWord;
const
  Buff_Size = MAX_COMPUTERNAME_LENGTH + 1;
begin
  nSize := Buff_Size;
  lpBuffer := StrAlloc(Buff_Size);
  GetComputerName(lpBuffer, nSize);
  Result := string(lpBuffer);
  StrDispose(lpBuffer);
end;

function TFrmPrincipal.validaTransportadora(codigoTransportadora: string; var codiEmpresa: string): Boolean;
begin
  if codigoTransportadora <> '' then
  begin
    qryConsulta.close;
    qryConsulta.SQL.Text := 'select cd_transportadora ,cd_empresa from T_TRANSPORTADORA WHERE cd_transportadora = ' + quotedstr(codigoTransportadora);
    qryConsulta.Open;

    if not qryConsulta.Eof then
    begin
      codiEmpresa := qryConsulta.fieldByname('cd_empresa').asstring;
      Result := True;
    end
    else
    begin
      codiEmpresa := '';
      Result := False;
    end;

  end
  else
  begin
    codiEmpresa := '';
    Result := False;
  end;

end;

end.

