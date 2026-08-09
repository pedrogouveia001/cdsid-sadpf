unit untEntrada;

interface

uses
  {Classes, SysUtils, IWAppForm, IWApplication, IWColor, IWTypes, pngimage,
  IWVCLBaseControl, IWBaseControl, IWBaseHTMLControl, IWControl, IWExtCtrls,
  Controls, Forms, IWVCLBaseContainer, IWContainer, IWHTMLContainer,
  IWHTML40Container, IWRegion, jpeg, IWCompLabel, IWCompRectangle, IWCompButton,
  IWCompText, IWHTMLControls, IWCompEdit, IWTMSCtrls,
  IWAdvToolButton;}

  Classes, SysUtils, IWAppForm, IWApplication, IWColor, IWTypes, IWCompCheckbox,
  IWExtCtrls, IWCompEdit, IWCompListbox, IWTMSCtrls, IWCompText, IWControl,
  IWWebGrid, IWAdvWebGrid, pngimage, jpeg, IWCompButton, IWVCLBaseControl,
  IWBaseControl, IWBaseHTMLControl, IWCompLabel, Controls, Forms,
  IWVCLBaseContainer, IWContainer, IWHTMLContainer, IWHTML40Container, IWRegion,
  IWHTMLControls, TeeGDIPlus, TeEngine, Series, ExtCtrls, TeeProcs, Chart,

  IWCompMemo, ComObj, Variants, windows, ActiveX,
  IWAdvWebGridExcel, Menus, IWCompMenu, IWCompTabControl, Math,
  IWBaseComponent,
  IWBaseHTML40Component, dateutils, IdExplicitTLSClientServerBase,
  IdMessageClient, IdSMTPBase, IdSMTP,
  IWCompRectangle,
  IWGrids, ActnList, IWDBStdCtrls, WebComp, WebAdapt, IdMessage,
  IdTCPConnection,
  IdTCPClient, Graphics,
  IdBaseComponent, IdComponent, IdIOHandler, IdIOHandlerSocket,
  IdIOHandlerStack, IdSSL, IdSSLOpenSSL, IdText, IWProgressIndicator,
  IWCompProgressBar, IWBaseHTMLComponent,
  IWTMSPopup, IWHelpTip, IWAdvToolButton;
type
  TtfrmEntrada = class(TIWAppForm)
    rgncorpo: TIWRegion;
    rgnbordaexterna: TIWRegion;
    rgnbordainterna: TIWRegion;
    rgnborda: TIWRegion;
    IWRectangle3: TIWRectangle;
    IWImage29: TIWImage;
    IWImage28: TIWImage;
    rgnInicio: TIWRegion;
    IWImage23: TIWImage;
    IWLabel3: TIWLabel;
    btncadastrar: TIWButton;
    btnentrar: TIWButton;
    rgnlogo: TIWRegion;
    IWRectangle1: TIWRectangle;
    IWRectangle2: TIWRectangle;
    IWLabel1: TIWLabel;
    IWLabel2: TIWLabel;
    IWImage1: TIWImage;
    rgnlogin: TIWRegion;
    rgnBElogin: TIWRegion;
    rgnBIlogin: TIWRegion;
    IWRegion4: TIWRegion;
    IWRectangle6: TIWRectangle;
    IWImage4: TIWImage;
    IWImage5: TIWImage;
    IWRegion5: TIWRegion;
    IWRectangle7: TIWRectangle;
    IWRectangle8: TIWRectangle;
    IWLabel4: TIWLabel;
    IWLabel5: TIWLabel;
    IWImage6: TIWImage;
    lbluserid: TIWLabel;
    edtemail: TIWEdit;
    lblSenha: TIWLabel;
    edtSenha: TIWEdit;
    IWLabel8: TIWLabel;
    IWLabel9: TIWLabel;
    lnkesqueceusenha: TIWLink;
    btnEnter: TIWButton;
    IWButton2: TIWButton;
    IWLink1: TIWLink;
    IWLabel6: TIWLabel;
    IWRegion1: TIWRegion;
    IWRectangle4: TIWRectangle;
    IWImage2: TIWImage;
    IWImage3: TIWImage;
    TIWHTMLLabel1: TTIWHTMLLabel;
    rgncadastro: TIWRegion;
    rgnBECadastro: TIWRegion;
    rgnBICadastro: TIWRegion;
    IWRegion7: TIWRegion;
    IWRectangle5: TIWRectangle;
    IWImage7: TIWImage;
    IWImage8: TIWImage;
    IWRegion8: TIWRegion;
    IWRectangle9: TIWRectangle;
    IWRectangle10: TIWRectangle;
    IWLabel7: TIWLabel;
    IWLabel10: TIWLabel;
    IWImage9: TIWImage;
    IWText2: TIWText;
    IWLabel104: TIWLabel;
    IWLabel105: TIWLabel;
    edtcadname: TIWEdit;
    IWLabel115: TIWLabel;
    IWLabel114: TIWLabel;
    edtcadlast: TIWEdit;
    IWLabel107: TIWLabel;
    IWLabel106: TIWLabel;
    edtcadmatricula: TIWEdit;
    IWLabel109: TIWLabel;
    IWLabel108: TIWLabel;
    edtcadestado: TIWEdit;
    IWLabel11: TIWLabel;
    IWLabel12: TIWLabel;
    edtcademail: TIWEdit;
    IWLabel13: TIWLabel;
    IWLabel14: TIWLabel;
    edtcadsenha: TIWEdit;
    IWLabel15: TIWLabel;
    IWLabel19: TIWLabel;
    edtcadconfirmarsenha: TIWEdit;
    IWLabel20: TIWLabel;
    btncadusuario: TIWButton;
    lnkvoltar: TIWLink;
    TIWAdvToolButton1: TTIWAdvToolButton;
    rgnnovasenha: TIWRegion;
    IWLabel17: TIWLabel;
    edtnovasenha: TIWEdit;
    IWLabel18: TIWLabel;
    edtconfirmarnovasenha: TIWEdit;
    IWButton3: TIWButton;
    IWButton4: TIWButton;
    txtpassrecover: TIWText;
    IWLabel16: TIWLabel;
    IWLabel21: TIWLabel;
    IWLabel22: TIWLabel;
    IWLabel23: TIWLabel;
    IWLink29: TIWLink;
    IWLink3: TIWLink;
    IWLink4: TIWLink;
    procedure rgncorpoCreate(Sender: TObject);
    procedure rgnInicioCreate(Sender: TObject);
    procedure rgnloginCreate(Sender: TObject);
    procedure btnentrarClick(Sender: TObject);
    procedure IWButton2Click(Sender: TObject);
    procedure btncadastrarClick(Sender: TObject);
    procedure IWLink1Click(Sender: TObject);
    procedure rgncadastroCreate(Sender: TObject);
    procedure lnkvoltarClick(Sender: TObject);
    procedure btnEnterClick(Sender: TObject);
    procedure btncadusuarioClick(Sender: TObject);
    procedure lnkesqueceusenhaClick(Sender: TObject);
    function send_email(destinatario: string; assunto: string;
  corpoMensagem: string): boolean;
    procedure IWButton3Click(Sender: TObject);
    procedure IWButton4Click(Sender: TObject);
    procedure IWLink29Click(Sender: TObject);
  public
  end;

implementation

uses untResultados, UserSessionUnit, uCode, ServerController;

{$R *.dfm}


function TtfrmEntrada.send_email(destinatario: string; assunto: string;
  corpoMensagem: string): boolean;
var
  // variáveis e objetos necessários para o envio
  IdSSLIOHandlerSocket: TIdSSLIOHandlerSocketOpenSSL;
  IdSMTP: TIdSMTP;
  IdMessage: TIdMessage;
  IdText: TIdText;
  sAnexo: string;
begin
  // Result:=False;

  // instanciação dos objetos
  IdSSLIOHandlerSocket := TIdSSLIOHandlerSocketOpenSSL.Create(Self);
  IdSMTP := TIdSMTP.Create(Self);
  IdMessage := TIdMessage.Create(Self);

  try
    // Configuração do protocolo SSL (TIdSSLIOHandlerSocketOpenSSL)
    IdSSLIOHandlerSocket.SSLOptions.Method := sslvSSLv23;
    IdSSLIOHandlerSocket.SSLOptions.Mode := sslmClient;
    IdSSLIOHandlerSocket.ReadTimeout := 30000;

    // Configuração do servidor SMTP (TIdSMTP)
    // Configuração do servidor SMTP (TIdSMTP)
    IdSMTP.IOHandler := IdSSLIOHandlerSocket;
    IdSMTP.UseTLS := utUseImplicitTLS;
    IdSMTP.AuthType := satDefault;
    IdSMTP.Port := 465;
    IdSMTP.Host := 'smtp.gmail.com';
    IdSMTP.Username := 'sadpf@cdsid.org.br';
    IdSMTP.Password := 'sadpf.cdsid';
    // IdSMTP.ReadTimeout:=30000;

    // Configuração da mensagem (TIdMessage)
    IdMessage.From.Address := 'sadpf@cdsid.org.br';
    IdMessage.From.Name := 'SADPF - SUPORTE';
    IdMessage.ReplyTo.EMailAddresses := IdMessage.From.Address;
    IdMessage.Recipients.EMailAddresses := destinatario;
    IdMessage.Subject := assunto;
    IdMessage.Body.Clear;
    IdMessage.Body.Add(corpoMensagem);

    // Conexão e autenticação
    try
      IdSMTP.Connect;
      IdSMTP.Authenticate;
    except
      on E: Exception do
      begin
        // WebApplication.ShowMessage('Erro na conexão ou autenticação: '+
        // MessageDlg('Erro na conexão ou autenticação: ' +
        // E.Message);
        Result := false;
        Exit;
      end;
    end;

    // Envio da mensagem
    try
      IdSMTP.Send(IdMessage);
      Result := true;
      // MessageDlg('Mensagem enviada com sucesso!', mtInformation, [mbOK], 0);
    except
      On E: Exception do
      begin
        // MessageDlg('Erro ao enviar a mensagem: ' +
        // E.Message, mtWarning, [mbOK], 0);
        Result := false;
      end;
    end;
  finally
    // desconecta do servidor
    IdSMTP.Disconnect;
    // liberação da DLL
    // UnLoadOpenSSLLibrary;
    // liberação dos objetos da memória
    FreeAndNil(IdMessage);
    FreeAndNil(IdSSLIOHandlerSocket);
    FreeAndNil(IdSMTP);
  end;

end;



procedure TtfrmEntrada.btncadastrarClick(Sender: TObject);
begin
rgninicio.visible := False;
rgncadastro.Visible := True;
end;

procedure TtfrmEntrada.btnEnterClick(Sender: TObject);
// procedimento de login no sistema
var
  email, id_email: string;
  senha, validation: string;
  quantregis: Integer;
  nome: string;

begin
  IWLabel9.Visible := false;
  IWLabel8.Visible := false;
  TIWUserSession(WebApplication.Data).email := edtemail.Text;
 // TIWUserSession(WebApplication.Data).primeiroacesso := false;
  // inicialmente fazendo um selectcount para verificar se já existe algum cadastro com esse email
  with UserSession.zquery2 do
  begin

    SQL.Text := 'SELECT count(*) FROM usuario WHERE email="' + (edtemail.Text)
      + '"';
    ExecSQL;
    open;
    quantregis := FieldByName('count(*)').AsInteger;
    close;
  end;
  if quantregis = 0 then
  begin
    IWLabel9.Visible := true;
    // IWButton1.Visible := true;
    Exit;
  end
  else
  begin
    TIWUserSession(WebApplication.Data).email := edtemail.Text;
    with UserSession.zquery2 do
    begin



      SQL.Text := 'SELECT senha FROM usuario WHERE email ="' + (edtemail.Text)
        + '"';
      ExecSQL;
      open;
      senha := FieldByName('senha').AsString;
      close;

    end;
    if senha = edtSenha.Text then
    begin
      with UserSession.zquery2 do
      begin
        SQL.Text := 'SELECT * FROM usuario WHERE email="' + (edtemail.Text) + '"';
        ExecSQL;
        open;
        TIWUserSession(WebApplication.Data).ID_user := FieldByName('Id_user')
          .AsString;

        TIWUserSession(WebApplication.Data).nome := FieldByName('nome')
          .AsString;

        TIWUserSession(WebApplication.Data).email := edtemail.Text;
        TIWUserSession(WebApplication.Data).senha := edtSenha.Text;
        if FieldByName('validado').AsString = 'Sim' then
        ttfrmResultados.Create(WebApplication).Show
        else
        rgnnovasenha.Visible := true;
        close;
      end;
    end
    else
      IWLabel8.Visible := true;
  end;
end;

procedure TtfrmEntrada.btnentrarClick(Sender: TObject);
begin
rgninicio.Visible := False;
rgnlogin.Visible := True;
end;

procedure TtfrmEntrada.btncadusuarioClick(Sender: TObject);
var
genero, statuscurso, nescolaridade, P, email: string;
  vetchar: array [0 .. 9] of string;
  i, a: Integer;
  H: boolean;
  Auxnomesobrenome: string;
  Auxdatanascimento: string;
  AuxsalvarBD: Integer;
begin
if (edtcadname.Text = '') or
(edtcadlast.Text = '') or
(edtcadmatricula.Text = '') or
(edtcadestado.Text = '') or
(edtcademail.Text = '') or
(edtcadsenha.Text = '') or
(edtcadconfirmarsenha.Text = '') then
begin
  WebApplication.ShowMessage('Por favor, preencha todos os campos.');
    Exit;
end;
  if edtcadsenha.Text <> edtcadconfirmarsenha.Text then
  begin
    WebApplication.ShowMessage('As senhas devem ser iguais!');
    Exit;
  end;
  vetchar[0] := '!';
  vetchar[1] := '@';
  vetchar[2] := '#';
  vetchar[3] := '$';
  vetchar[4] := '%';
  vetchar[5] := '&';
  vetchar[6] := '*';
  vetchar[7] := '^';
  vetchar[8] := '-';
  vetchar[9] := '+';

  email := edtcademail.Text;
  a := 0;
  for i := 1 to length(email) do
  begin
    if (email[i] = '@') or (email[i] = '.') then
      INC(a);
  end;
  if a < 2 then
  begin
    WebApplication.ShowMessage(
      'Por favor, introduza um endereço de e-mail válido.');
    Exit;
  end;

  TIWUserSession(WebApplication.Data).ID_user := '';
  with UserSession.zquery2 do
  begin
    SQL.Text := 'SELECT Id_user FROM usuario WHERE email ="' + (edtcademail.Text)
      + '";';
    ExecSQL;
    open;
    TIWUserSession(WebApplication.Data).ID_user := FieldByName('Id_user').AsString;
    close;
  end;

  // * se o e-mail já estiver cadastrado *//
  if (TIWUserSession(WebApplication.Data).ID_user <> '') then
  begin
    WebApplication.ShowMessage(
      'O endereço de e-mail informado já foi cadastrado. Caso não se lembre da sua senha, é possível criar um novo cadastro.');
    Exit;
  end;

   With TiwUserSession(WebApplication.Data) do
   begin
   with UserSession.zquery2 do
  begin
    // salvando as informações contidas na região cadusuario, email, titulo, primeiro nome, sobrenome, tipo de filiação, senha registrada, cidade,pais, nome da instituição, curso, qual o proposito para usar fitradeoff,data de nascimento, nível de escolaridade, status, genero

    // ESSE CASO FOI CHAMADO NO BOTÃO btnSend (SEND) DA REGIÃO rgnCadusuario (QUANDO SE FAZ O CADASTRO DE UM NOVO USUARIO)
    // como todas as informações já estão no formato de string basta utilizar o comando para salvar
    // não estou entendendo porque o ; esta marcado como errado
    // inicialmente juntando o nome e o sobrenome
    Auxnomesobrenome := edtcadname.Text + ' ' + edtcadlast.Text;
    SQL.Text :=
      'INSERT INTO usuario (nome, email, senha, estado, matricula, prob_finalizado,validado) VALUES ("' + edtcadname.text + '","' + edtcademail.Text + '","' + edtcadsenha.Text + '", "' + edtcadestado.Text + '", "' + edtcadmatricula.Text + '", "' + 'Não' + '", "' + 'Sim' + '")';
    ExecSQL;

    SQL.Text := 'SELECT Id_user FROM usuario WHERE email="' + (edtcademail.Text) + '"';
        ExecSQL;
        open;
        TIWUserSession(WebApplication.Data).ID_user := FieldByName('Id_user')
          .AsString;
        close;



    // CRITÉRIO DANO A PESSOA HUMANA
    SQL.Text :=
      'INSERT INTO criterios (Id_user, nome, prefdir, niveis) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + 'Dano à pessoa humana' + '","' + 'maximização' + '", "' + '6' + '")';
    ExecSQL;


    SetLength(Id_criterio, 9);
    SQL.Text := 'SELECT Id_criterio FROM criterios WHERE Id_user="' + ID_user + '" AND nome = "' + 'Dano à pessoa humana'+ '"';
        ExecSQL;
        open;
        Id_Criterio[0] := FieldByName('Id_criterio').AsString;
        close;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[0] + '","' + 'Dano à pessoa humana' + '","' + '0' + '", "' + 'Que não atente contra a vida, honra, saúde etc de ninguém diretamente' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[0] + '","' + 'Dano à pessoa humana' + '","' + '1' + '", "' + 'Que atente contra a honra ou saúde, integridade física e/ou mental' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[0] + '","' + 'Dano à pessoa humana' + '","' + '2' + '", "' + 'Que atente contra a honra, saúde e integridade física, ainda que de uma pessoa, sendo ela vulnerável ou servidor público ou agente do Estado em razão das funções por ele desempenhadas' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[0] + '","' + 'Dano à pessoa humana' + '","' + '3' + '", "' + 'Que atente contra a vida de uma pessoa' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[0] + '","' + 'Dano à pessoa humana' + '","' + '4' + '", "' + 'Que atente contra a vida de uma ou mais pessoas em atividade típica de grupo de extermínio/ Que atente contra a vida de servidor público ou agente do Estado em razão das funções por ele desempenhadas' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[0] + '","' + 'Dano à pessoa humana' + '","' + '5' + '", "' + 'Que configure genocídio' + '")';
    ExecSQL;


   // CRITERIO DANO AO MEIO AMBIENTE
   SQL.Text :=
      'INSERT INTO criterios (Id_user, nome, prefdir, niveis) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + 'Dano ao Meio Ambiente' + '","' + 'maximização' + '", "' + '6' + '")';
    ExecSQL;


    SetLength(Id_criterio, 9);
    SQL.Text := 'SELECT Id_criterio FROM criterios WHERE Id_user="' + ID_user + '" AND nome = "' + 'Dano ao Meio Ambiente'+ '"';
        ExecSQL;
        open;
        Id_Criterio[1] := FieldByName('Id_criterio')
          .AsString;
        close;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[1] + '","' + 'Dano ao Meio Ambiente' + '","' + '0' + '", "' + 'Que não cause dano ou risco ao meio ambiente ou à saúde pública' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[1] + '","' + 'Dano ao Meio Ambiente' + '","' + '1' + '", "' + 'Com potencialidade de causar um impacto ambiental' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[1] + '","' + 'Dano ao Meio Ambiente' + '","' + '2' + '", "' + 'Que ponha em risco a saúde pública' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[1] + '","' + 'Dano ao Meio Ambiente' + '","' + '3' + '", "' + 'Que atinja a saúde pública' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[1] + '","' + 'Dano ao Meio Ambiente' + '","' + '4' + '", "' + 'Que provoque um impacto ambiental de médias proporções' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[1] + '","' + 'Dano ao Meio Ambiente' + '","' + '5' + '", "' + 'Que provoque um impacto ambiental de grandes proporções' + '")';
    ExecSQL;

    // CRITERIO Reprovabilidade da Conduta
   SQL.Text :=
      'INSERT INTO criterios (Id_user, nome, prefdir, niveis) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + 'Reprovabilidade da Conduta' + '","' + 'maximização' + '", "' + '6' + '")';
    ExecSQL;


    SetLength(Id_criterio, 9);
    SQL.Text := 'SELECT Id_criterio FROM criterios WHERE Id_user="' + ID_user + '" AND nome = "' + 'Reprovabilidade da Conduta'+ '"';
        ExecSQL;
        open;
        Id_Criterio[2] := FieldByName('Id_criterio')
          .AsString;
        close;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[2] + '","' + 'Reprovabilidade da Conduta' + '","' + '0' + '", "' + 'Que admita transação penal' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[2] + '","' + 'Reprovabilidade da Conduta' + '","' + '1' + '", "' + 'Até 05 anos' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[2] + '","' + 'Reprovabilidade da Conduta' + '","' + '2' + '", "' + 'Até 10 anos' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[2] + '","' + 'Reprovabilidade da Conduta' + '","' + '3' + '", "' + 'Entre 10 a 15 anos' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[2] + '","' + 'Reprovabilidade da Conduta' + '","' + '4' + '", "' + 'Entre 15 a 20 anos' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[2] + '","' + 'Reprovabilidade da Conduta' + '","' + '5' + '", "' + 'Mais de 20 anos' + '")';
    ExecSQL;

    // Repercussão e Extensão do Crime
   SQL.Text :=
      'INSERT INTO criterios (Id_user, nome, prefdir, niveis, somatorio) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + 'Repercussão e Extensão do Crime' + '","' + 'maximização' + '", "' + '6' + '", "' + 'Sim' + '")';
    ExecSQL;


    SetLength(Id_criterio, 9);
    SQL.Text := 'SELECT Id_criterio FROM criterios WHERE Id_user="' + ID_user + '" AND nome = "' + 'Repercussão e Extensão do Crime'+ '"';
        ExecSQL;
        open;
        Id_Criterio[3] := FieldByName('Id_criterio')
          .AsString;
        close;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[3] + '","' + 'Repercussão e Extensão do Crime' + '","' + '0' + '", "' + 'Prática de corrupção e/ou outros crimes por servidores públicos ou agentes políticos' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[3] + '","' + 'Repercussão e Extensão do Crime' + '","' + '1' + '", "' + 'Tráfico de drogas' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[3] + '","' + 'Repercussão e Extensão do Crime' + '","' + '2' + '", "' + 'Lavagem de dinheiro' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[3] + '","' + 'Repercussão e Extensão do Crime' + '","' + '3' + '", "' + 'Crimes hediondos ou equiparados' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[3] + '","' + 'Repercussão e Extensão do Crime' + '","' + '4' + '", "' + 'Com a participação de organização criminosa' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[3] + '","' + 'Repercussão e Extensão do Crime' + '","' + '5' + '", "' + 'Recursos públicos desviados' + '")';
    ExecSQL;

   // CRITERIO Proximidade da Prescrição
   SQL.Text :=
      'INSERT INTO criterios (Id_user, nome, prefdir, niveis) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + 'Proximidade da Prescrição' + '","' + 'maximização' + '", "' + '6' + '")';
    ExecSQL;


    SetLength(Id_criterio, 9);
    SQL.Text := 'SELECT Id_criterio FROM criterios WHERE Id_user="' + ID_user + '" AND nome = "' + 'Proximidade da Prescrição'+ '"';
        ExecSQL;
        open;
        Id_Criterio[4] := FieldByName('Id_criterio')
          .AsString;
        close;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[4] + '","' + 'Proximidade da Prescrição' + '","' + '0' + '", "' + 'Fatos sujeitos à prescrição em 20 anos ou imprescritíveis' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[4] + '","' + 'Proximidade da Prescrição' + '","' + '1' + '", "' + 'Fatos sujeitos à prescrição em mais de seis anos' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[4] + '","' + 'Proximidade da Prescrição' + '","' + '2' + '", "' + 'Fatos sujeitos à prescrição em até seis anos' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[4] + '","' + 'Proximidade da Prescrição' + '","' + '3' + '", "' + 'Fatos sujeitos à prescrição em até quatro anos' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[4] + '","' + 'Proximidade da Prescrição' + '","' + '4' + '", "' + 'Fatos sujeitos à prescrição em até dois anos' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[4] + '","' + 'Proximidade da Prescrição' + '","' + '5' + '", "' + 'Fatos sujeitos à prescrição em até um ano' + '")';
    ExecSQL;

    // CRITERIO Estágio da Investigação
   SQL.Text :=
      'INSERT INTO criterios (Id_user, nome, prefdir, niveis) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + 'Estágio da Investigação' + '","' + 'maximização' + '", "' + '6' + '")';
    ExecSQL;


    SetLength(Id_criterio, 9);
    SQL.Text := 'SELECT Id_criterio FROM criterios WHERE Id_user="' + ID_user + '" AND nome = "' + 'Estágio da Investigação'+ '"';
        ExecSQL;
        open;
        Id_Criterio[5] := FieldByName('Id_criterio')
          .AsString;
        close;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[5] + '","' + 'Estágio da Investigação' + '","' + '0' + '", "' + 'Realizadas diligências ou não no âmbito policial sem expectativa da necessidade de cautelares' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[5] + '","' + 'Estágio da Investigação' + '","' + '1' + '", "' + 'Realizadas diligências ou não no âmbito policial sem ainda a representação de cautelares, embora já seja vislumbrada a necessidade de tais medidas' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[5] + '","' + 'Estágio da Investigação' + '","' + '2' + '", "' + 'Realizadas diligências no âmbito policial com representação policial de cautelares' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[5] + '","' + 'Estágio da Investigação' + '","' + '3' + '", "' + 'Realizadas diligências no âmbito policial, com deferimento judicial de cautelares' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[5] + '","' + 'Estágio da Investigação' + '","' + '4' + '", "' + 'Realizadas diligências no âmbito policial com cautelares implementadas em caso de interceptação telefônica ou telemática e, no caso de mandados de busca, com endereços e alvos devidamente identificados e mapeados' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[5] + '","' + 'Estágio da Investigação' + '","' + '5' + '", "' + 'Realizadas diligências no âmbito policial com cautelares implementadas em caso de interceptação telefônica ou telemática e, no caso de mandados de prisão, com endereços e alvos devidamente identificados e mapeados' + '")';
    ExecSQL;



    // CRITERIO Necessidade de Recursos Humanos especializados
   SQL.Text :=
      'INSERT INTO criterios (Id_user, nome, prefdir, niveis) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + 'Necessidade de Recursos Humanos especializados' + '","' + 'maximização' + '", "' + '6' + '")';
    ExecSQL;


    SetLength(Id_criterio, 9);
    SQL.Text := 'SELECT Id_criterio FROM criterios WHERE Id_user="' + ID_user + '" AND nome = "' + 'Necessidade de Recursos Humanos especializados'+ '"';
        ExecSQL;
        open;
        Id_Criterio[6] := FieldByName('Id_criterio')
          .AsString;
        close;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[6] + '","' + 'Necessidade de Recursos Humanos especializados' + '","' + '0' + '", "' + 'Sem indicativo de necessidade de diárias' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[6] + '","' + 'Necessidade de Recursos Humanos especializados' + '","' + '1' + '", "' + 'Diárias de dois a quatro meses' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[6] + '","' + 'Necessidade de Recursos Humanos especializados' + '","' + '2' + '", "' + 'Diárias de quatro a seis meses ou com previsão de deflagração unicamente com recursos locais' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[6] + '","' + 'Necessidade de Recursos Humanos especializados' + '","' + '3' + '", "' + 'Diárias de seis a oito meses ou com previsão de deflagração com o apoio de unidades policiais dentro do mesmo Estado' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[6] + '","' + 'Necessidade de Recursos Humanos especializados' + '","' + '4' + '", "' + 'Diárias de oito a dez meses ou previsão de deflagração com o apoio de unidades policiais em Estados próximos, com deslocamento por viatura' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[6] + '","' + 'Necessidade de Recursos Humanos especializados' + '","' + '5' + '", "' + 'Diárias por mais de dez meses ou previsão de deflagração com o apoio de unidades policiais de outros Estados, com deslocamento aéreo ou, por qualquer razão, com necessidade de uso de aeronave' + '")';
    ExecSQL;


    // CRITERIO Estimativa de valor de bens apreendidos ou sequestrados
   SQL.Text :=
      'INSERT INTO criterios (Id_user, nome, prefdir, niveis) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + 'Estimativa de valor de bens apreendidos ou sequestrados' + '","' + 'maximização' + '", "' + '6' + '")';
    ExecSQL;


    SetLength(Id_criterio, 9);
    SQL.Text := 'SELECT Id_criterio FROM criterios WHERE Id_user="' + ID_user + '" AND nome = "' + 'Estimativa de valor de bens apreendidos ou sequestrados'+ '"';
        ExecSQL;
        open;
        Id_Criterio[7] := FieldByName('Id_criterio')
          .AsString;
        close;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[7] + '","' + 'Estimativa de valor de bens apreendidos ou sequestrados' + '","' + '0' + '", "' + 'Operações especiais sem previsão de bens apreendidos ou sequestrados ou com estimativa de apreensão ou sequestro em até 5 milhões de reais' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[7] + '","' + 'Estimativa de valor de bens apreendidos ou sequestrados' + '","' + '1' + '", "' + 'Operações especiais com previsão de bens apreendidos ou sequestrados, em valores de 5 a 10 milhões' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[7] + '","' + 'Estimativa de valor de bens apreendidos ou sequestrados' + '","' + '2' + '", "' + 'Operações especiais com previsão de bens apreendidos ou sequestrados, em valores de 10 a 15 milhões' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[7] + '","' + 'Estimativa de valor de bens apreendidos ou sequestrados' + '","' + '3' + '", "' + 'Operações especiais com previsão de bens apreendidos ou sequestrados, em valores de 15 a 20 milhões' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[7] + '","' + 'Estimativa de valor de bens apreendidos ou sequestrados' + '","' + '4' + '", "' + 'Operações especiais com previsão de bens apreendidos ou sequestrados, em valores de 20 a 25 milhões' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[7] + '","' + 'Estimativa de valor de bens apreendidos ou sequestrados' + '","' + '5' + '", "' + 'Operações especiais com previsão de bens apreendidos ou sequestrados, em valores acima de 25 milhões' + '")';
    ExecSQL;

    // CRITERIO Estimativa de prejuízos evitados à União
   SQL.Text :=
      'INSERT INTO criterios (Id_user, nome, prefdir, niveis) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + 'Estimativa de prejuízos evitados à União' + '","' + 'maximização' + '", "' + '6' + '")';
    ExecSQL;


    SetLength(Id_criterio, 9);
    SQL.Text := 'SELECT Id_criterio FROM criterios WHERE Id_user="' + ID_user + '" AND nome = "' + 'Estimativa de prejuízos evitados à União'+ '"';
        ExecSQL;
        open;
        Id_Criterio[8] := FieldByName('Id_criterio')
          .AsString;
        close;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[8] + '","' + 'Estimativa de prejuízos evitados à União' + '","' + '0' + '", "' + 'Operações especiais cuja deflagração comporte a estimativa de evitar prejuízo à União na ordem de até 5 milhões nos 12 meses seguintes ou cujo objeto investigativo já está exaurido ou não tenha aspecto econômico/financeiro' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[8] + '","' + 'Estimativa de prejuízos evitados à União' + '","' + '1' + '", "' + 'Operações especiais cuja deflagração comporte a estimativa de evitar prejuízo à União na ordem de 5 a 10 milhões nos 12 meses seguintes' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[8] + '","' + 'Estimativa de prejuízos evitados à União' + '","' + '2' + '", "' + 'Operações especiais cuja deflagração comporte a estimativa de evitar prejuízo à União na ordem de 10 a 15 milhões nos 12 meses seguintes' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[8] + '","' + 'Estimativa de prejuízos evitados à União' + '","' + '3' + '", "' + 'Operações especiais cuja deflagração comporte a estimativa de evitar prejuízo à União na ordem de 15 a 20 milhões nos 12 meses seguintes' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[8] + '","' + 'Estimativa de prejuízos evitados à União' + '","' + '4' + '", "' + 'Operações especiais cuja deflagração comporte a estimativa de evitar prejuízo à União na ordem de 20 a 25 milhões nos 12 meses seguintes' + '")';
    ExecSQL;

    SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[8] + '","' + 'Estimativa de prejuízos evitados à União' + '","' + '5' + '", "' + 'Operações especiais cuja deflagração comporte a estimativa de evitar prejuízo à União na ordem de mais de 25 milhões nos 12 meses seguintes' + '")';
    ExecSQL;



  end;
   end;

  WebApplication.ShowMessage(
      'Cadastro bem sucedido. Você já pode acessar o sistema.');


rgncadastro.Visible := False;
rgnlogin.Visible := True;
end;

procedure TtfrmEntrada.IWButton2Click(Sender: TObject);
begin
rgninicio.Visible := true;
rgnlogin.Visible := false;
end;

procedure TtfrmEntrada.IWButton3Click(Sender: TObject);
begin
with UserSession.ZQuery2 do
    begin
      if (edtnovasenha.Text <> '') and (edtconfirmarnovasenha.Text = edtnovasenha.Text) then
      begin
      // salvando no BD, o campo validacao atual
      SQL.Text := 'UPDATE usuario SET validado ="Sim" WHERE email="' +
        (TiwUserSession(WebApplication.Data).email) + '";';
      ExecSQL;

      // salvando no BD,a senha temporária
      SQL.Text := 'UPDATE usuario SET senha ="' +
        edtnovasenha.Text
        + '" WHERE email="' + (TiwUserSession(WebApplication.Data).email) + '";';
      ExecSQL;
      rgnnovasenha.Visible := false;
      end;
    end;
end;

procedure TtfrmEntrada.IWButton4Click(Sender: TObject);
begin
rgnnovasenha.Visible := false;
end;

procedure TtfrmEntrada.IWLink1Click(Sender: TObject);
begin
rgnlogin.visible := False;
rgncadastro.Visible := True;
end;

procedure TtfrmEntrada.IWLink29Click(Sender: TObject);
begin
WebApplication.SendFile('C:\inetpub\SADSDS\Files\Guia Rápido do Usuário - SADPF' +
      '.pdf', true);
end;


procedure TtfrmEntrada.lnkvoltarClick(Sender: TObject);
begin
rgncadastro.Visible := False;
rgnlogin.Visible := True;
end;



procedure TtfrmEntrada.lnkesqueceusenhaClick(Sender: TObject);
var
  P, nome, titulo: string;
  vetchar: array [0 .. 9] of string;
  i, a, quantregis: integer;
  H: boolean;
begin
  TiwUserSession(WebApplication.Data).email := edtemail.Text;
  if TiwUserSession(WebApplication.Data).email = '' then
  begin
    WebApplication.ShowMessage(
      'Por favor, preencha o email e tente novamente.');
    Exit;
  end;
  with UserSession.ZQuery2 do
  begin

    SQL.Text := 'SELECT count(*) FROM usuario WHERE email="' + (edtemail.Text)
      + '"';
    ExecSQL;
    open;
    quantregis := FieldByName('count(*)').AsInteger;
    close;
  end;
  if quantregis = 0 then
  begin
    IWLabel9.Visible := true;
    Exit;
  end
  else
    IWLabel9.Visible := false;

  vetchar[0] := '!';
  vetchar[1] := '@';
  vetchar[2] := '#';
  vetchar[3] := '$';
  vetchar[4] := '%';
  vetchar[5] := '&';
  vetchar[6] := '*';
  vetchar[7] := '^';
  vetchar[8] := '-';
  vetchar[9] := '+';

  with UserSession.ZQuery2 do
  begin
    SQL.Text := 'SELECT nome FROM usuario WHERE email ="' + (edtemail.Text) + '"';
    ExecSQL;
    open;
    nome := FieldByName('nome').AsString;
    close;


  end;

  for i := 0 to 5 - 1 do
  begin
    P := FloatToStr(RandomRange(0, 9));
    TIWUserSession(WebApplication.Data).auxvalidacao := TIWUserSession
      (WebApplication.Data).auxvalidacao + FloatToStr
      (Round(RandomRange(0, 1000) / 10)) + vetchar[StrToInt(P)];

  end;


    txtpassrecover.Lines[0] := txtpassrecover.Lines[0] + ' ' + nome + ',';
  txtpassrecover.Lines[5] := txtpassrecover.Lines[5] + ' ' + TIWUserSession
    (WebApplication.Data).auxvalidacao;

  H := false;
  H := send_email(TiwUserSession(WebApplication.Data).email, 'SADPF - Senha Temporária',
    txtpassrecover.Lines.Text);

  if H then
  begin
    WebApplication.ShowMessage(
      ' Foi enviada uma senha temporária para o e-mail informado. Deverá alterá-la no seu próximo acesso ao SADPF.'
        + ' Verifique também o seu spam, caso a mensagem não chegue à sua caixa de entrada.');
    with UserSession.ZQuery2 do
    begin
      // salvando no BD, o campo validacao atual
      SQL.Text := 'UPDATE usuario SET validado ="senha" WHERE email="' +
        (TiwUserSession(WebApplication.Data).email) + '";';
      ExecSQL;

      // salvando no BD,a senha temporária
      SQL.Text := 'UPDATE usuario SET senha ="' +
        (TIWUserSession(WebApplication.Data).auxvalidacao)
        + '" WHERE email="' + (TiwUserSession(WebApplication.Data).email) + '";';
      ExecSQL;
    end;
  end
  else
    WebApplication.ShowMessage(
      'Não foi possível concluir esta ação. Por favor, tente novamente mais tarde. Se o problema persistir, contate-nos.');
  txtpassrecover.Lines[0] := 'Prezado(a)';
  txtpassrecover.Lines[5] := 'Senha Temporária:';
end;

procedure TtfrmEntrada.rgncadastroCreate(Sender: TObject);
begin
rgncadastro.Align := AlClient;
end;

procedure TtfrmEntrada.rgncorpoCreate(Sender: TObject);
begin
rgncorpo.Align := AlClient;
end;

procedure TtfrmEntrada.rgnInicioCreate(Sender: TObject);
begin
rgninicio.Align := AlClient;
end;

procedure TtfrmEntrada.rgnloginCreate(Sender: TObject);
begin
rgnlogin.Align := AlClient;
end;

initialization
  TtfrmEntrada.SetAsMainForm;

end.
