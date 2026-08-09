unit untResultados;

interface

uses

 { Classes, SysUtils, IWAppForm, IWApplication, IWColor, IWTypes, IWCompEdit,
  IWVCLBaseControl, IWBaseControl, IWBaseHTMLControl, IWControl, IWCompLabel,
  Controls, Forms, IWVCLBaseContainer, IWContainer, IWHTMLContainer,
  IWHTML40Container, IWRegion, TeeGDIPlus, IWCompButton, TeEngine, ExtCtrls,
  TeeProcs, Chart, IWExtCtrls, IWWebGrid, IWAdvWebGrid, BubbleCh, TeePolar,
  IWCompListbox, IWCompRadioButton, IWTMSCheckList, series, IWGrids, TeeLisB,
  TeeSeriesTextEd, Buttons, Math, pngimage, CheckLst,
  ShellApi, TeeLegendScrollBar, Spin, ToolWin, StdActns,
  TeeURL, TeeXML, AppEvnts, IWCompDynamicChart, IWVCLComponent, TeeProCo,
  Windows, Messages, Variants, Graphics, Dialogs, ComCtrls, Menus, StdCtrls,
  jpeg, ComObj, Grids, IWCompText, IWHTMLControls, ErrorBar,
  IWAdvWebGridExcel, IWProgressIndicator, IWDBGrids,
  IWCompProgressBar, IWBaseComponent, IWBaseHTMLComponent,
  IWBaseHTML40Component, IWColorComboBox, IWTMSImgCtrls, IWTMSCtrls,
  IWCompTabControl, IWTMSPopup, IWTMSMenus, RpDefine, RpRender, RpRenderText,
  IWAdvCheckGroup, IWDBAdvCheckGroup, IWHelpTip, IWCompCheckbox, GanttCh,
  IwCompRectangle,
  IWCompMemo, IWAdvToolButton, ActnList; }

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
  IWTMSPopup, IWHelpTip,IWAdvToolButton, IWCompRadioButton,
  IWTMSCheckList, ToolWin, IWTMSImgCtrls,

  TeeSeriesTextEd, Spin, StdActns,
  TeeURL, TeeXML, AppEvnts, IWCompDynamicChart, IWVCLComponent, TeeProCo,
  Messages, Dialogs, ComCtrls, StdCtrls;
type
  TtfrmResultados = class(TIWAppForm)
    rgnopcoes: TIWRegion;
    rbnBEOpcoes: TIWRegion;
    rgnBIOpcoes: TIWRegion;
    rgnborda: TIWRegion;
    IWRectangle3: TIWRectangle;
    IWImage29: TIWImage;
    IWImage28: TIWImage;
    rgnlogo: TIWRegion;
    IWRectangle1: TIWRectangle;
    IWRectangle2: TIWRectangle;
    IWLabel1: TIWLabel;
    IWLabel2: TIWLabel;
    IWImage1: TIWImage;
    IWText2: TIWText;
    IWLabel3: TIWLabel;
    IWLabel4: TIWLabel;
    TIWHTMLLabel1: TTIWHTMLLabel;
    btnEnter: TIWButton;
    IWLink2: TIWLink;
    rdbop: TIWRadioButton;
    rdbcrit: TIWRadioButton;
    rgnoperacoes: TIWRegion;
    rgnBEOperacoes: TIWRegion;
    rgnBIOperacoes: TIWRegion;
    IWRegion1: TIWRegion;
    IWRectangle4: TIWRectangle;
    IWImage5: TIWImage;
    IWImage6: TIWImage;
    IWRegion2: TIWRegion;
    IWRectangle5: TIWRectangle;
    IWRectangle6: TIWRectangle;
    IWLabel5: TIWLabel;
    IWLabel6: TIWLabel;
    IWImage7: TIWImage;
    btnVisualizarOP: TTIWAdvToolButton;
    btnCadastrarOP: TTIWAdvToolButton;
    rgnvisualizarOP: TIWRegion;
    IWRegion4: TIWRegion;
    grdroperacoes: TTIWAdvWebGrid;
    IWLink1: TIWLink;
    lnkexcluirOP: TIWLink;
    lnkalterarOP: TIWLink;
    rgncriarOP: TIWRegion;
    IWLabel7: TIWLabel;
    edtnomeOP: TIWEdit;
    IWLabel8: TIWLabel;
    IWLabel9: TIWLabel;
    mmDetalhesOP: TIWMemo;
    btnProsseguir: TIWButton;
    rgncriterios: TIWRegion;
    rgnBEcriterios: TIWRegion;
    rgnBIcriterios: TIWRegion;
    IWRegion3: TIWRegion;
    IWRectangle7: TIWRectangle;
    IWImage8: TIWImage;
    IWImage9: TIWImage;
    IWRegion5: TIWRegion;
    IWRectangle8: TIWRectangle;
    IWRectangle9: TIWRectangle;
    IWLabel10: TIWLabel;
    IWLabel11: TIWLabel;
    IWImage10: TIWImage;
    IWLink3: TIWLink;
    rgnvisualizarcriterios: TIWRegion;
    IWLabel12: TIWLabel;
    cmbbxCriterios: TIWComboBox;
    lblprefdir: TIWLabel;
    IWLabel13: TIWLabel;
    rgnniveiscriterios: TIWRegion;
    lbldirecaopreferencia: TIWLabel;
    lblnumerodeniveis: TIWLabel;
    IWButton2: TIWButton;
    btnexcluircriterio: TIWButton;
    IWLink4: TIWLink;
    rgncriarnovocriterio: TIWRegion;
    IWLabel15: TIWLabel;
    IWLabel16: TIWLabel;
    IWLabel17: TIWLabel;
    btncriarCriterio: TIWButton;
    rgnpreencherniveis: TIWRegion;
    IWLabel20: TIWLabel;
    IWRegion9: TIWRegion;
    grdpreencherniveis: TTIWAdvWebGrid;
    edtnomecrit: TIWEdit;
    cmbbxprefdir: TIWComboBox;
    cmbbxniveiscriterios: TIWComboBox;
    edtn1: TIWEdit;
    edtn2: TIWEdit;
    edtn3: TIWEdit;
    edtn4: TIWEdit;
    edtn5: TIWEdit;
    edtn6: TIWEdit;
    edtn7: TIWEdit;
    rgnquestionario: TIWRegion;
    rgnBEQuestionario: TIWRegion;
    rgnBIQuestionario: TIWRegion;
    IWRegion7: TIWRegion;
    IWRectangle10: TIWRectangle;
    IWImage11: TIWImage;
    IWImage12: TIWImage;
    IWRegion8: TIWRegion;
    IWRectangle11: TIWRectangle;
    IWRectangle12: TIWRectangle;
    IWLabel18: TIWLabel;
    IWLabel19: TIWLabel;
    IWImage13: TIWImage;
    IWLink5: TIWLink;
    IWLabel21: TIWLabel;
    rgnpergunta: TIWRegion;
    btnsalvarquestionario: TIWButton;
    lbloperacaosel1: TIWLabel;
    lbllabel24: TIWLabel;
    rdbpriori: TIWRadioButton;
    rgnresultados: TIWRegion;
    rgnberesultado: TIWRegion;
    rgnbiresultado: TIWRegion;
    IWRegion13: TIWRegion;
    IWRectangle16: TIWRectangle;
    IWImage17: TIWImage;
    IWImage18: TIWImage;
    IWRegion14: TIWRegion;
    IWRectangle17: TIWRectangle;
    IWRectangle18: TIWRectangle;
    IWLabel26: TIWLabel;
    IWLabel27: TIWLabel;
    IWImage19: TIWImage;
    IWLink7: TIWLink;
    IWLabel28: TIWLabel;
    IWLink8: TIWLink;
    rgnelicitacao: TIWRegion;
    rgnBEelicitacao: TIWRegion;
    rgnBIelicitacao: TIWRegion;
    IWRegion15: TIWRegion;
    IWRectangle19: TIWRectangle;
    IWImage20: TIWImage;
    IWImage21: TIWImage;
    IWRegion16: TIWRegion;
    IWRectangle20: TIWRectangle;
    IWRectangle21: TIWRectangle;
    IWLabel30: TIWLabel;
    IWLabel31: TIWLabel;
    IWImage22: TIWImage;
    lstbxcriterios: TIWListbox;
    btnreiniciar: TIWButton;
    btnselecionar: TIWButton;
    IWLabel34: TIWLabel;
    lstbxordemcrit: TIWListbox;
    btnverresultados: TIWButton;
    IWLink12: TIWLink;
    grdDetalhamentoCrit: TIWGrid;
    IWLabel35: TIWLabel;
    IWText3: TIWText;
    btnsalvarop: TIWButton;
    clsbtxCriterioEspecial: TTIWCheckListBox;
    rgnprocroc: TIWRegion;
    btngoroc: TIWButton;
    IWRegion29: TIWRegion;
    lblNSOE: TIWLabel;
    txtS2OE: TIWText;
    IWRegion30: TIWRegion;
    lstbxrankROC: TIWListbox;
    lstbxcriterioROC: TIWListbox;
    lblCOSCOOE: TIWLabel;
    Btnrestart: TIWButton;
    btnchooseROC: TIWButton;
    IWRegion17: TIWRegion;
    imgGrafOverallE: TIWImage;
    btngerargraficoavhol: TIWButton;
    lblS1OE: TIWLabel;
    IWRegion31: TIWRegion;
    lblBiOE: TIWLabel;
    lblWiOE: TIWLabel;
    crtOverallE: TChart;
    Series4: TBarSeries;
    Series1: TBarSeries;
    IWLink11: TIWLink;
    btnvisualizacaotabular: TTIWAdvToolButton;
    btnvisualizacaografica: TTIWAdvToolButton;
    IWLabel39: TIWLabel;
    IWLabel14: TIWLabel;
    grdcriterios: TTIWAdvWebGrid;
    btneditarcrit: TIWButton;
    grdexportacao: TTIWAdvWebGrid;
    TIWAdvWebGridExcelIO1: TTIWAdvWebGridExcelIO;
    IWLink13: TIWLink;
    rgnimgop: TIWRegion;
    rgnimgcrit: TIWRegion;
    rgnimgpriori: TIWRegion;
    imgpriori: TIWImage;
    imgcrit: TIWImage;
    imgop: TIWImage;
    edtdia: TIWEdit;
    edtmes: TIWEdit;
    edtano: TIWEdit;
    IWLabel24: TIWLabel;
    IWLabel25: TIWLabel;
    IWLabel32: TIWLabel;
    IWLink14: TIWLink;
    IWLink15: TIWLink;
    IWLink16: TIWLink;
    IWLink17: TIWLink;
    IWLink18: TIWLink;
    IWLink19: TIWLink;
    rgninputmanual: TIWRegion;
    rgnbeinputmanual: TIWRegion;
    rgnbiinputmanual: TIWRegion;
    IWRegion10: TIWRegion;
    IWRectangle13: TIWRectangle;
    IWImage2: TIWImage;
    IWImage3: TIWImage;
    IWRegion11: TIWRegion;
    IWRectangle14: TIWRectangle;
    IWRectangle15: TIWRectangle;
    IWLabel40: TIWLabel;
    IWLabel41: TIWLabel;
    IWImage4: TIWImage;
    IWLink20: TIWLink;
    IWLink21: TIWLink;
    IWLabel42: TIWLabel;
    IWText1: TIWText;
    rgnavaliacaodireta: TIWRegion;
    IWText4: TIWText;
    IWRegion20: TIWRegion;
    grdinputmanual: TIWGrid;
    btnsalvar: TIWButton;
    btnredefiniravdireta: TIWButton;
    btnverresultadosavdireta: TIWButton;
    IWLink22: TIWLink;
    btnhasse: TTIWAdvToolButton;
    btntabela: TTIWAdvToolButton;
    rgnresultadotabular: TIWRegion;
    IWLabel22: TIWLabel;
    rgnordemop: TIWRegion;
    grdresultado: TIWGrid;
    IWLabel29: TIWLabel;
    IWRegion6: TIWRegion;
    txtordemcrit: TIWText;
    IWLink10: TIWLink;
    rgnresultadohasse: TIWRegion;
    TIWFadeImage2: TTIWFadeImage;
    tmrhasse: TIWTimer;
    btncriarcrit: TIWButton;
    TIWHelpTip1: TTIWHelpTip;
    TIWHelpTip2: TTIWHelpTip;
    IWLabel43: TIWLabel;
    IWText5: TIWText;
    IWLabel33: TIWLabel;
    edtdiafato: TIWEdit;
    IWLabel44: TIWLabel;
    IWLabel45: TIWLabel;
    edtmesfato: TIWEdit;
    IWLabel46: TIWLabel;
    edtanofato: TIWEdit;
    IWLink23: TIWLink;
    IWLink24: TIWLink;
    IWLink25: TIWLink;
    IWLink26: TIWLink;
    IWLink27: TIWLink;
    IWLink28: TIWLink;
    IWLink29: TIWLink;
    IWLabel47: TIWLabel;
    IWLabel48: TIWLabel;
    IWLabel49: TIWLabel;
    IWLabel50: TIWLabel;
    IWLabel51: TIWLabel;
    IWLabel52: TIWLabel;
    IWLabel53: TIWLabel;
    rgnquestoescrit: TIWRegion;
    rgnbequestoescrit: TIWRegion;
    rgnbiquestoescrit: TIWRegion;
    IWLink33: TIWLink;
    IWLabel55: TIWLabel;
    btnsalvarquestoescrit: TIWButton;
    lblcritsel1: TIWLabel;
    IWLabel61: TIWLabel;
    IWLink34: TIWLink;
    IWRegion23: TIWRegion;
    IWRegion26: TIWRegion;
    IWRectangle25: TIWRectangle;
    IWImage23: TIWImage;
    IWImage24: TIWImage;
    IWRegion27: TIWRegion;
    IWRectangle26: TIWRectangle;
    IWRectangle27: TIWRectangle;
    IWLabel62: TIWLabel;
    IWLabel63: TIWLabel;
    IWImage25: TIWImage;
    IWLink35: TIWLink;
    IWLabel64: TIWLabel;
    grdquestionariocrit: TTIWAdvWebGrid;
    rgniveis: TIWRegion;
    grdniveis: TTIWAdvWebGrid;
    procedure rgnopcoesCreate(Sender: TObject);
    procedure btnEnterClick(Sender: TObject);
    procedure rgnoperacoesCreate(Sender: TObject);
    procedure btnVisualizarOPClick(Sender: TObject);
    procedure btnCadastrarOPClick(Sender: TObject);
    procedure cmbbxCriteriosChange(Sender: TObject);
    procedure rgncriteriosCreate(Sender: TObject);
    procedure cmbbxniveiscriteriosChange(Sender: TObject);
    procedure IWLink4Click(Sender: TObject);
    procedure IWLink3Click(Sender: TObject);
    procedure rgnquestionarioCreate(Sender: TObject);
    procedure btnProsseguirClick(Sender: TObject);
    procedure rgnresultadosCreate(Sender: TObject);
    procedure rgnelicitacaoCreate(Sender: TObject);
    procedure lstbxcriteriosChange(Sender: TObject);
    procedure btnselecionarClick(Sender: TObject);
    procedure IWAppFormCreate(Sender: TObject);
    procedure btnsalvaropClick(Sender: TObject);
    procedure IWLink1Click(Sender: TObject);
    procedure btncriarCriterioClick(Sender: TObject);
    procedure CriarQuestionario;
    procedure btnsalvarquestionarioClick(Sender: TObject);
    procedure lnkalterarOPClick(Sender: TObject);
    procedure lnkexcluirOPClick(Sender: TObject);
    procedure DeletarQuestionario;
    procedure IWLink5Click(Sender: TObject);
    procedure btnexcluircriterioClick(Sender: TObject);
    procedure btnvisualizacaotabularClick(Sender: TObject);
    procedure btnvisualizacaograficaClick(Sender: TObject);
    procedure btnreiniciarClick(Sender: TObject);
    procedure CalculaRoc;
    procedure Smarter;
    procedure Normaliza;
    procedure gerar_grafico_overall;
    procedure Intra_Razao;
    procedure btnverresultadosClick(Sender: TObject);
    Function casas_decimais(Num: Double): String;
    procedure lstbxcriterioROCClick(Sender: TObject);
    procedure btnchooseROCClick(Sender: TObject);
    procedure BtnrestartClick(Sender: TObject);
    procedure btngorocClick(Sender: TObject);
    procedure IWLink10Click(Sender: TObject);
    procedure IWButton2Click(Sender: TObject);
    procedure btneditarcritClick(Sender: TObject);
    procedure IWLink11Click(Sender: TObject);
    procedure IWLink7Click(Sender: TObject);
    procedure IWLink8Click(Sender: TObject);
    procedure IWLink13Click(Sender: TObject);
    procedure imgopClick(Sender: TObject);
    procedure imgcritClick(Sender: TObject);
    procedure imgprioriClick(Sender: TObject);
    procedure IWLink14Click(Sender: TObject);
    procedure CriarElementosAvDireta;
    procedure DestruirElementosAvDireta;
    procedure btnsalvarClick(Sender: TObject);
    procedure btnredefiniravdiretaClick(Sender: TObject);
    procedure IWLink12Click(Sender: TObject);
    procedure IWLink20Click(Sender: TObject);
    procedure rgninputmanualCreate(Sender: TObject);
    procedure btnverresultadosavdiretaClick(Sender: TObject);
    procedure tmrhasseTimer(Sender: TObject);
    procedure btnhasseClick(Sender: TObject);
    procedure btntabelaClick(Sender: TObject);
    procedure btncriarcritClick(Sender: TObject);
    function TrocaPV(str : string) : string;
    function TrocaVP(str : string) : string;
    procedure IWLink23Click(Sender: TObject);
    procedure rgnquestoescritCreate(Sender: TObject);
    procedure IWLink33Click(Sender: TObject);
    procedure btnsalvarquestoescritClick(Sender: TObject);

  public
  end;

implementation

uses untEntrada, UserSessionUnit, ServerController;
{$R *.dfm}

function TtfrmResultados.TrocaPV(str : string): string;
var
i : integer;
begin
  for i := 1 to length(str) do
  begin
    if str[i] = '.' then
    str[i] := ',';
  end;
  Result := str;
end;

function TtfrmResultados.TrocaVP(str : string): string;
var
i : integer;
begin
  for i := 1 to length(str) do
  begin
    if str[i] = ',' then
    str[i] := '.';
  end;
  Result := str;
end;

Function TtfrmResultados.casas_decimais(Num: Double): String;
var
  S: string;
  i, k: integer;
  inicia: Boolean;
begin
  if Num <= 10000 then
  begin
  // ** O número de casas decimais está definido para 2 **//
  inicia := false;
  k := 0;
  S := FloatToStr(Num);
  // ** Contando o número de casas decimais **//
  for i := 2 to Length(S) - 1 do
  begin
    if inicia = true then
      INC(k);
    if (S[i] = ',') or (S[i] = '.') then
      inicia := true;
  end;
  // ** Só precisa formatar se o número de casas decimais for maior que 2 **//
  if k >= 3 then
    Result := FormatFloat('0.0000', Num)
  else
    Result := S;
  end
  else
  begin
   k := 0;
   S := FloatToStr(Num);
   // ** Contando o número de casas decimais **//
  for i := 0 to Length(S) - 1 do
  begin
  if (S[i] = ',') or (S[i] = '.') then
  break;

  INC(k);
  end;
  Num := Num / Power(10,k);
  S := FormatFloat('0.0000', Num);
  Num := StrToFloat(S) * Power(10,k);
  Result := FloatToStr(num);
  end;
end;

procedure TtfrmResultados.Intra_Razao;
var
  i: integer;
begin
  with TIWUserSession(WebApplication.Data) do
  begin
    // ** Setando o tamanho dos vetores **//
    SetLength(valor_razao, Numcrit);
    SetLength(BConseqOrd, Numcrit);
    SetLength(WConseqOrd, Numcrit);

    // ** Preenchendo vetores **//
    for i := 0 to Numcrit - 1 do
      // ** Critério de maximização **//
      if (Prefdir[i] = 1) then
      begin
        BConseqOrd[i] := MaxConseq[i];
        WConseqOrd[i] := minConseq[i];
      end
      else
      begin
        // ** Critério de minimização **//
       { BConseqOrd[i] := minConseq[i];
        WConseqOrd[i] := MaxConseq[i]; }
        BConseqOrd[i] := MaxConseq[i];
        WConseqOrd[i] := minConseq[i];
      end;

    for i := 0 to Numcrit - 1 do
    begin

      if WConseqOrd[i] > BConseqOrd[i] then
      /// condição de minimização ///
      begin
        if (WConseqOrd[i] < 0) and (BConseqOrd[i] < 0) then
        // fórmula para caso negativo
        begin
          valor_razao[i] := (WConseqOrd[i] / BConseqOrd[i]);
        end;
        if (WConseqOrd[i] > 0) and (BConseqOrd[i] > 0) then
        // fórmula para caso positivo
        begin
          valor_razao[i] := (BConseqOrd[i] / WConseqOrd[i]);
        end;
        if (WConseqOrd[i] < 0) and (BConseqOrd[i] > 0) then
        // fórmula para caso misto
        begin
          valor_razao[i] := ((WConseqOrd[i] / (BConseqOrd[i] - WConseqOrd[i]))
              + 1 - (BConseqOrd[i] / (BConseqOrd[i] - WConseqOrd[i])));
        end;
        if (BConseqOrd[i] < 0) and (WConseqOrd[i] > 0) then
        // fórmula para caso misto
        begin
          valor_razao[i] := ((WConseqOrd[i] / (BConseqOrd[i] - WConseqOrd[i]))
              + 1 - (BConseqOrd[i] / (BConseqOrd[i] - WConseqOrd[i])));
        end;
      end; // do meu primeiro if

      if WConseqOrd[i] < BConseqOrd[i] then
      /// condição de maximização///
      begin
        if (WConseqOrd[i] < 0) and (BConseqOrd[i] < 0) then
        // fórmula para caso negativo
        begin
          valor_razao[i] := BConseqOrd[i] / WConseqOrd[i];
        end;
        if (WConseqOrd[i] >= 0) and (BConseqOrd[i] >= 0) then
        // fórmula para caso positivo
        begin
          valor_razao[i] := WConseqOrd[i] / BConseqOrd[i];
        end;
        if (WConseqOrd[i] < 0) and (BConseqOrd[i] > 0) then
        // fórmula para caso misto
        begin
          valor_razao[i] := (WConseqOrd[i] / (BConseqOrd[i] - WConseqOrd[i]))
            + 1 - (BConseqOrd[i] / (BConseqOrd[i] - WConseqOrd[i]));
        end;
        if (BConseqOrd[i] < 0) and (WConseqOrd[i] > 0) then
        begin
          valor_razao[i] := (WConseqOrd[i] / (BConseqOrd[i] - WConseqOrd[i]))
            + 1 - (BConseqOrd[i] / (BConseqOrd[i] - WConseqOrd[i]));
        end;
      end; // do 2º if
    end; // do for
  end; // do with
end;

procedure TtfrmResultados.gerar_grafico_overall;
var
  i: integer;
  bt1, wt1, Bt: string;
begin
  // ** Caso passe a ser bilingue, essas variáveis seriam mudadas **//
  bt1 := 'M';
  wt1 := 'P';
  Bt := '';
  Intra_Razao;
  imgGrafOverallE.Visible := true;
  // torna o componente IWImage visivel
  crtOverallE.BottomAxis.Visible := true; // torna o eixo x visível
  crtOverallE.SeriesList.ClearValues;
  // limpar as séries se ja houver algum dado
  crtOverallE.leftAxis.Items.Clear; // limpando os valores dos eixos para não se sobrepor

  with TIWUserSession(WebApplication.Data) do
  begin
    SetLength(vetorgrafico, Numcrit);
    crtOverallE.leftAxis.Items.Add(100, Bt); // adicionando o nome best no eixo y do melhor cirtério
    for i := 0 to Numcrit - 1 do
    begin
      crtOverallE.BottomAxis.Items.Add(i, 'Crit ' + IntToStr(i + 1));
      if vetorGrafico[i] = 1 then
      // ** O critério está selecionado no listbox  **//
      begin
        if Prefdir[i] = 1 then
        crtOverallE.Series[0].AddXY(i, 100, bt1 + IntToStr(i + 1)
            + ':' + casas_decimais(TIWUserSession(WebApplication.Data)
              .BConseqOrd[i]), clwebYELLOW)
        else
         crtOverallE.Series[0].AddXY(i, 100, bt1 + IntToStr(i + 1)
            + ':' + casas_decimais(TIWUserSession(WebApplication.Data)
              .WConseqOrd[i]), clwebYELLOW)
      end;

      if vetorGrafico[i] = 2 then
      // ** Critério ja ordenado **//
      begin
        if Prefdir[i] = 1 then
        crtOverallE.Series[0].AddXY(i, 100, bt1 + IntToStr(i + 1)
            + ':' + casas_decimais(TIWUserSession(WebApplication.Data)
              .BConseqOrd[i]), clWebGREEN)
        else
         crtOverallE.Series[0].AddXY(i, 100, bt1 + IntToStr(i + 1)
            + ':' + casas_decimais(TIWUserSession(WebApplication.Data)
              .WConseqOrd[i]), clWebGREEN)
      end;
      if vetorGrafico[i] = 0 then
      begin
        // ** O critério ainda não foi nem selecionado, nem ordenado **//
        if Prefdir[i] = 1 then
        begin
        crtOverallE.Series[0].AddXY(i, 100, bt1 + IntToStr(i + 1)
            + ':' + casas_decimais(TIWUserSession(WebApplication.Data).BConseqOrd[i]));

        crtOverallE.Series[1].AddXY(i, (WConseqOrd[i]/BConseqOrd[i])*100, wt1 + IntToStr(i + 1)
            + ':' + casas_decimais(TIWUserSession(WebApplication.Data).WConseqOrd[i]), clWebRed);

        end
        else
        begin
        crtOverallE.Series[0].AddXY(i, 100, bt1 + IntToStr(i + 1)
            + ':' + casas_decimais(TIWUserSession(WebApplication.Data)
              .WConseqOrd[i]));

        crtOverallE.Series[1].AddXY(i, (WConseqOrd[i]/BConseqOrd[i])*100, wt1 + IntToStr(i + 1)
            + ':' + casas_decimais(TIWUserSession(WebApplication.Data).BConseqOrd[i]), clWebRed);

        end;


        {if i = 0 then
        crtOverallE.Series[1].AddXY(i, TIWUserSession(WebApplication.Data)
            .valor_razao[i] * 100, wt1 + IntToStr(i + 1) + ':' + casas_decimais
            (TIWUserSession(WebApplication.Data).BConseqOrd[i]), clWebRED)
        else
        crtOverallE.Series[1].AddXY(i, TIWUserSession(WebApplication.Data)
            .valor_razao[i] * 90, wt1 + IntToStr(i + 1) + ':' + casas_decimais
            (TIWUserSession(WebApplication.Data).WConseqOrd[i]), clWebRED) }
      end;
    end;

    imgGrafOverallE.Picture.Bitmap.Assign // comando que tira o 'print' do gráfico para exibi-lo no componente image
      (crtOverallE.TeeCreateBitmap(clWebAQUA, rect(0, 0, imgGrafOverallE.Width,
          imgGrafOverallE.Height)));
  end;
end;

procedure TtfrmResultados.CriarElementosAvDireta;
var
i, j, k : integer;
begin
  With TiwUserSession(WebApplication.Data) do
  begin
    setlength(lblcritavdireta, numcrit);
    setlength(edtcritavdireta, numcrit);
    for j := 0 to Numcrit - 1 do
    begin
    lblcritavdireta[j] := TIWLabel.Create(rgnavaliacaodireta);
    with lblcritavdireta[j] do
    begin
      Parent := rgnavaliacaodireta;
      Left := 35 + 85 * j;
      top := 10;
      Caption := 'Crit ' + inttostr(j+1);
      Font.FontName := 'Verdana';
      Font.Color := clWebGAINSBORO;
      //Name := 'lblCelColunaPeso' + IntToStr(j);
      // lbl.OnAsyncKeyPress:=IWEdit1.onas
    end;
    edtcritavdireta[j] := TIWEdit.Create(rgnavaliacaodireta);
    with edtcritavdireta[j] do
    begin
      //Name := 'edtCelPeso' + IntToStr(j);
      Height := 21;
      Width := 54;
      Left := 25 + 85 * j;
      top := 32;
      Parent := rgnavaliacaodireta;
      Font.Color := clblack;
      Text := '';
      if (btngoroc.Enabled = true) or (btnverresultados.enabled = true) then
      Text := FormatFloat('0.00',TiwUserSession(WebApplication.Data).Peso[j]);

    end;
    end;

    grdinputmanual.Clear;
    grdinputmanual.ColumnCount := Numcrit + 1;
    grdinputmanual.RowCount := 3;
    grdinputmanual.Cell[0,0].Text := 'Critério:';
    grdinputmanual.Cell[1,0].Text := 'Sigla:';
    grdinputmanual.Cell[2,0].Text := 'Posição:';
    for j := 0 to grdinputmanual.RowCount - 1 do
    grdinputmanual.Cell[j,0].font.Style := [fsbold];
    for I := 0 to Numcrit - 1 do
    begin
    grdinputmanual.Cell[0,1+i].Text := Nomecrit[i];
    grdinputmanual.Cell[1,1+i].Text := 'Crit ' + inttostr(i+1);
    end;
  end;


end;

procedure TtfrmResultados.DestruirElementosAvDireta;
var
i, j, k : integer;
begin
  With TiwUserSession(WebApplication.Data) do
  begin
  for i := 0 to Numcrit - 1 do
  begin
    lblcritavdireta[i].Destroy;
    edtcritavdireta[i].Destroy;
  end;
  end;
end;



procedure TtfrmResultados.imgcritClick(Sender: TObject);
begin
rgnimgcrit.BorderOptions.Color := clWebGOLDENROD;
rgnimgcrit.BorderOptions.NumericWidth := 2;

rgnimgop.BorderOptions.Color := clnone;
//rgnimgop.BorderOptions.NumericWidth := 0;

rgnimgpriori.BorderOptions.Color := clnone;
//rgnimgpriori.BorderOptions.NumericWidth := 0;

rdbop.Checked := False;
rdbpriori.Checked := False;
rdbcrit.Checked := True;
end;

procedure TtfrmResultados.imgopClick(Sender: TObject);
begin
rgnimgop.BorderOptions.Color := clWebGOLDENROD;
rgnimgop.BorderOptions.NumericWidth := 2;

rgnimgpriori.BorderOptions.Color := clnone;
//rgnimgpriori.BorderOptions.NumericWidth := 0;

rgnimgcrit.BorderOptions.Color := clnone;
//rgnimgcrit.BorderOptions.NumericWidth := 0;

rdbop.Checked := True;
rdbpriori.Checked := False;
rdbcrit.Checked := False;

end;

procedure TtfrmResultados.imgprioriClick(Sender: TObject);
begin
rgnimgpriori.BorderOptions.Color := clWebGOLDENROD;
rgnimgpriori.BorderOptions.NumericWidth := 2;

rgnimgop.BorderOptions.Color := clnone;
//rgnimgop.BorderOptions.NumericWidth := 0;

rgnimgcrit.BorderOptions.Color := clnone;
//rgnimgcrit.BorderOptions.NumericWidth := 0;

rdbop.Checked := false;
rdbpriori.Checked := true;
rdbcrit.Checked := False;

end;

procedure TtfrmResultados.DeletarQuestionario;
var
i,j,k : integer;
begin
  With TiwUserSession(WebApplication.Data) do
  begin
    for I := 0 to TiwUserSession(WebApplication.Data).Numcrit - 1 do
    begin
      txtperguntas[i].Destroy;
      rdgniveis[i].Destroy;
    end;
  end;
end;



procedure TtfrmResultados.CalculaROC;
 var
  i, j, k: integer;
  auxpeso, soma: real;
begin
  for i := 0 to TIWUserSession(WebApplication.Data).Numcrit - 1 do
  begin
    soma := 0;
    for j := i + 1 to TIWUserSession(WebApplication.Data).Numcrit do
    begin
      soma := soma + 1 / j;
    end;
    auxpeso := (1 / TIWUserSession(WebApplication.Data).Numcrit) * soma;
    TIWUserSession(WebApplication.Data).peso[TIWUserSession(WebApplication.Data).ordemroc[i]] := auxpeso;

  end;

   SetLength(TIWUserSession(WebApplication.Data).ordemcrit, TIWUserSession(WebApplication.Data).numcrit);
   for i := 0 to TIWUserSession(WebApplication.Data).Numcrit - 1 do
        begin
          k := 0;
          for j := 0 to TIWUserSession(WebApplication.Data).Numcrit - 1 do
          begin
            if i = j then
            continue;
            if TIWUserSession(WebApplication.Data).peso[i] < TIWUserSession(WebApplication.Data).peso[j] then
            k := k + 1;
          end;
          TIWUserSession(WebApplication.Data).ordemcrit[i] := k;
        end;



  for i := 0 to TIWUserSession(WebApplication.Data).Numcrit - 1 do
  begin
   With UserSession.zquery2 do
      begin
      SQL.Text := 'UPDATE criterios SET peso ="' + floattostr(TIWUserSession(WebApplication.Data).peso[i]) + '" WHERE Id_criterio=' + TIWUserSession(WebApplication.Data).Id_criterio[i];
      ExecSQL;
      end;
  end;
end;

procedure TtfrmResultados.Smarter;
var
i, j, Rlevels : integer;
pular : array of boolean;
auxniveis : array of real;
DomMat : string;
encontrado : boolean;
begin
 With TiwUserSession(WebApplication.Data) do
  begin



    Normaliza;
    finalize(resultadoop);
    finalize(rankingfinal);
    finalize(matrizparapar);
    SetLength(resultadoop, numop);
    SetLength(pular, numop);
    SetLength(RankingFinal, NumOP);
    SetLength(MatrizParAPar, NumOP, NUmOP);
    for j := 0 to NumOp - 1 do
    begin
      Pular[j] := false;
      for I := 0 to Numcrit - 1 do
      begin
        TiwUserSession(WebApplication.Data).ResultadoOP[j] := ResultadoOP[j] +
        TiwUserSession(WebApplication.Data).MatrizConseqNorm[j,i]*
        TiwUserSession(WebApplication.Data).Peso[i];
      end;
    end;

    NumNiveis := NumOP;
    for i := 0 to NumOP - 1 do
    begin
      for j := 0 to NUmOP - 1 do
      begin
        if i <> j then
        begin
          if ResultadoOP[i] > ResultadoOP[j] then
          begin
          MatrizParaPar[i,j] := 1;
          end;
          if ResultadoOP[i] < ResultadoOP[j] then
          MatrizParapar[i,j] := 0;
          if ResultadoOP[i] = ResultadoOP[j] then
          begin
          MatrizParapar[i,j] := 2;
          end;

        end;
        DomMat := DomMat + inttostr(MatrizParapar[i,j]);
      end;
    end;
    setlength(auxniveis,1);
    auxniveis[0] := ResultadoOP[0];
    for I := 1 to NumOP - 1 do
    begin
    encontrado := false;
      for j := 0 to length(auxniveis) - 1 do
      begin
        if ResultadoOP[i] = auxniveis[j] then
        begin
        encontrado := true;
        break;
        end;
      end;
      if encontrado = false then
      begin
        setlength(auxniveis, length(auxniveis)+1);
        auxniveis[length(auxniveis)-1] := ResultadoOP[i];
      end;
    end;

    RLevels := length(auxniveis);



     auxNomeOP := '';
    for I := 0 to NumOP - 1 do
      begin
        auxNomeOP := auxNomeOP + '§' + NomeOP[i];
      end;

    // Teste unindo numAlt e Levels
      levels_numalt := IntToStr(Rlevels) + '/' + IntToStr(NumOP);

      // Tirando o usuário da fila
      UserSession.ZQuery2.SQL.Text :=
        'DELETE FROM line_dh WHERE id_problem = "' + id_user + '"';
      ZQuery2.ExecSQL;

      UserSession.ZQuery2.SQL.Text :=
        'INSERT INTO line_dh (id_problem, id_user, dom_matrix, ah_matrix, NomeAlt, levels_numalt ) VALUES ("' + id_user + '", "' + ID_user + '", "' + DomMat + '", "' + '' + '", "' + auxNomeOP + '", "' + levels_numalt + '")';
      UserSession.ZQuery2.ExecSQL;

      UserSession.ZQuery2.SQL.Text :=
        'UPDATE desk SET dh_in_progress = "' + IntToStr(1) + '"';
      UserSession.ZQuery2.ExecSQL;

      tmrhasse.Enabled := true;





    for I := 0 to NumOP - 1 do
    begin
      rankingfinal[i] := 1;
      for j := 0 to NumOP - 1 do
      begin
      if i = j then
      continue;
      if ResultadoOP[j] > ResultadoOP[i] then
      TiwUserSession(WebApplication.Data).rankingfinal[i] := rankingfinal[i] + 1;
      end;
      With UserSession.zquery2 do
      begin
      SQL.Text := 'UPDATE operacoes SET ranking ="' + inttostr(rankingfinal[i]) + '" WHERE Id =' + Id_Operacao[i];
      ExecSQL;
      end;
    end;
     With UserSession.zquery2 do
      begin
      SQL.Text := 'UPDATE usuario SET prob_finalizado ="' + 'Sim' + '" WHERE Id_user =' + Id_user;
      ExecSQL;
      prob_finalizado := true;
      end;
  end;
end;

procedure TtfrmResultados.tmrhasseTimer(Sender: TObject);
var
  prob: boolean;
  T1: Pwidechar;
begin
  prob := false;
  try
    //TIWFadeImage1.Picture.LoadFromFile('C:\inetpub\fitradeoff\images\' +
      //  (TIWUserSession(WebApplication.Data).id_problema) + '.jpg');
    TIWFadeImage2.Picture.LoadFromFile('C:\inetpub\sadsds\images\' +
        (TIWUserSession(WebApplication.Data).id_user) + '.jpg');

    rgnresultadohasse.Height := TIWUserSession(WebApplication.Data).NumOP*100 + 500;
    rgnbiresultado.height := rgnresultadohasse.Height + 100;

    //TIWFadeImage2.left := -300;

   DeleteFile(Pchar('C:\inetpub\sadsds\images\' + TIWUserSession
         (WebApplication.Data).id_user + '.jpg'));
  except
    prob := True;
  end;
  // plotar_max_diff;
  if prob = false then
  begin
   //WebApplication.ShowMessage('An unexpected error has occurred, please, use the update button to generate diagram once again.');
    //lblhassemessage.Visible := false;
    //TbCtrlvisualizationsChange(self);
    // btnhasse.Enabled := true;
    tmrhasse.Enabled := false;

  end;
end;


procedure TtfrmResultados.Normaliza;
var
i, j : integer;
begin
  With TiwUserSession(WebApplication.Data) do
  begin
    for j := 0 to NumOp - 1 do
    begin
      for I := 0 to Numcrit - 1 do
      begin
      if (MinConseq[i] = MaxConseq[i])  then
      begin
      MatrizConseqNorm[j,i] := 0;
      peso[i] := 0;
      continue;
      end;
      if prefdir[i] = 1 then
      MatrizConseqNorm[j,i] := (MatrizConseq[j,i] - MinConseq[i])/(MaxConseq[i] - MinConseq[i])
      else
      begin
      MatrizConseqNorm[j,i] := (MatrizConseq[j,i] - MaxConseq[i])/(-MaxConseq[i] + MinConseq[i]);
      end;
      end;
    end;

  end;
end;

procedure TtfrmResultados.CriarQuestionario;
var
i,j,k : integer;
begin
  With TiwUserSession(WebApplication.Data) do
  begin
    SetLength(rgnperguntas, Numcrit);
    SetLength(txtperguntas, Numcrit);
    SetLength(rdgniveis, Numcrit);
    clsbtxCriterioEspecial.visible := false;
    clsbtxCriterioEspecial.Items.Clear;
    for I := 0 to TiwUserSession(WebApplication.Data).Numcrit - 1 do
    begin
      {rgnperguntas[i] := TiwRegion.Create(self);

        rgnperguntas[i].Parent := rgnbiquestionario;
        rgnperguntas[i].left := 82;
        rgnperguntas[i].height := 276;
        rgnperguntas[i].Width := 983;
        rgnperguntas[i].color := ClWebBlack;
        rgnperguntas[i].BorderOptions.Color := ClWebGoldenRod;
        if i = 0 then
        TiwUserSession(WebApplication.Data).rgnperguntas[i].Top := 93
        else
        rgnperguntas[i].Top := 113 + 276*i;
        rgnperguntas[i].visible := true;
        rgnperguntas[i].Zindex := i; }

        txtperguntas[i] := TiwText.Create(self);
        txtperguntas[i].Parent := rgnpergunta;
        txtperguntas[i].left := 32;
        if i = 0 then
        txtperguntas[i].top := 11
        else
        txtperguntas[i].top := 11 + 20 + 276*i;
        txtperguntas[i].BgColor := clWebBlack;
        txtperguntas[i].Font.FontName := 'verdana';
        txtperguntas[i].Font.Color := ClWebGoldenRod;
        txtperguntas[i].Width := 881;
        txtperguntas[i].Height := 57;
        txtperguntas[i].Lines[0] := 'C' +inttostr(i+1) + ' - ' + Nomecrit[i] +':';
        txtperguntas[i].font.Style := [fsBold];
        txtperguntas[i].font.Size := 11;
        txtperguntas[i].visible := true;

        rdgniveis[i] := TiwRadioGroup.Create(self);
        rdgniveis[i].Parent := rgnpergunta;
        rdgniveis[i].left := 32;
        if i = 0 then
        rdgniveis[i].top := 54
        else
        rdgniveis[i].top := 54 + 20 + 276*i; ;
        rdgniveis[i].width := 881;
        rdgniveis[i].height := 167;
        rdgniveis[i].font.FontName := 'verdana';
        rdgniveis[i].font.Color := clwebgainsboro;
        rdgniveis[i].font.size := 11;
        for j := 0 to Niveis[i] - 1 do
        begin
          rdgniveis[i].items.Add(inttostr(j) + ': ' + Ndetalhados[i,j]);
        end;
        rdgniveis[i].visible := true;

        rdgniveis[i].ItemIndex := Round(MatrizConseq[OpSel, i]);


        if CritEspecial[i] = true then
        begin
          txtperguntas[i].Lines[0] := 'C' + inttostr(i+1) + ' - ' + Nomecrit[i] + ':';
          rdgniveis[i].visible := False;
          clsbtxCriterioEspecial.visible := true;
          clsbtxCriterioEspecial.Top := rdgniveis[i].top;
          clsbtxCriterioEspecial.left := rdgniveis[i].left;
          clsbtxCriterioEspecial.height := rdgniveis[i].height;
          clsbtxCriterioEspecial.width := rdgniveis[i].width;
          clsbtxCriterioEspecial.items.Clear;
          for j := 0 to Niveis[i] - 1 do
        begin
          clsbtxCriterioEspecial.items.Add(Ndetalhados[i,j]);
        end;
        end;

      end;
    rgnbiquestionario.height := 300 + 276*Numcrit;
    rgnpergunta.Height := 150 + 276*numcrit - rgnpergunta.top;
    btnsalvarquestionario.Top := 150 + 276*Numcrit - btnsalvarquestionario.height;
  end;

end;





procedure TtfrmResultados.btnCadastrarOPClick(Sender: TObject);
begin
  rgncriarOP.visible := True;
  rgnvisualizarOP.visible := false;
  if rgncriarOP.visible = True then
  begin
    btnVisualizarOP.BackColor := clWebGoldenRod;
    btnVisualizarOP.backcolorto := clwebPeachPuff;
    btnCadastrarOP.BackColor := clwebsilver;
    btnCadastrarOP.backcolorto := clwebgray;
  end
  else
  begin
    btnVisualizarOP.BackColor := clwebsilver;
    btnVisualizarOP.backcolorto := clwebgray;
    btnCadastrarOP.BackColor := clWebGoldenRod;
    btnCadastrarOP.backcolorto := clwebPeachPuff;
  end;

end;

procedure TtfrmResultados.btnchooseROCClick(Sender: TObject);
var
  i: integer;
begin
  TIWUserSession(WebApplication.Data).vetorGrafico[lstbxcriterioROC.ItemIndex]
    := 2;
  gerar_grafico_overall;
  with TIWUserSession(WebApplication.Data) do
  begin
    SetLength(ordemroc, Numcrit); // ajustar
    if (lstbxcriterioROC.ItemIndex <> -1) and
      (lstbxcriterioROC.Items[lstbxcriterioROC.ItemIndex] <> '') then
    begin
      lstbxrankROC.Items.Add(IntToStr(lstbxrankROC.Items.Count + 1) + '- ' + Nomecrit[lstbxcriterioROC.ItemIndex]
        );
      ordemroc[lstbxrankROC.Items.Count - 1] := lstbxcriterioROC.ItemIndex;
      lstbxcriterioROC.Items[lstbxcriterioROC.ItemIndex] := '';
      if lstbxrankROC.Items.Count = Numcrit - 1 then
      begin
        for i := 0 to Numcrit - 1 do
          if lstbxcriterioROC.Items[i] <> '' then
          begin
            //lstbxrankROC.Items.Add(lstbxcriterioROC.Items[i]);
            ordemroc[lstbxrankROC.Items.Count - 1] := i;
            lstbxcriterioROC.Items[i] := '';
            TIWUserSession(WebApplication.Data).vetorGrafico[i] := 2;
          end;
        btnchooseROC.Enabled := false;
        btngoroc.Enabled := true;
        gerar_grafico_overall;
        // WebApplication.ShowMessage('ROC procedure completed', smAlert);
        CalculaROC;
        lstbxrankROC.Items.Clear;
    for i := 0 to Numcrit - 1 do
    begin
      lstbxrankROC.Items.Add(inttostr(i+1) + '- ' + Nomecrit[ordemroc[i]]+ ' (Peso = ' +
      FormatFloat('0.00',TiwUserSession(WebApplication.Data).peso[
      TiwUserSession(WebApplication.Data).ordemroc[i]]) + ')');
    end;
        btngoroc.Enabled := true;
      end;
    end;
  end;

end;

procedure TtfrmResultados.btncriarcritClick(Sender: TObject);
begin
edtn1.Clear;
edtn2.Clear;
edtn3.Clear;
edtn4.Clear;
edtn5.Clear;
edtn6.Clear;
edtn7.Clear;
edtn1.visible := false;
edtn2.visible := false;
edtn3.visible := false;
edtn4.visible := false;
edtn5.visible := false;
edtn6.visible := false;
edtn7.visible := false;
TIWUserSession(WebApplication.Data).NovoCriterio := true;
cmbbxniveiscriteriosChange(self);
rgnvisualizarcriterios.Visible := false;
rgncriarnovocriterio.Visible := true;
btncriarCriterio.Visible := true;
btneditarcrit.Visible := false;
end;

procedure TtfrmResultados.btncriarCriterioClick(Sender: TObject);
var
i, j : integer;
begin
   With TIWUserSession(WebApplication.Data) do
   begin
    if edtnomecrit.Text = '' then
    begin
    WebApplication.ShowMessage('Preencha corretamente as informações!');
    Exit;
    end;
    for i := 0 to Numcrit - 1 do
    begin
      if edtnomecrit.Text = nomecrit[i] then
      begin
        WebApplication.ShowMessage('Já existe um critério com este nome. Preencha novamente o campo!');
        Exit;
      end;
    end;
    Numcrit := Numcrit + 1;
    SetLength(NomeCrit,Numcrit);
    SetLength(Niveis,Numcrit);
    SetLength(NDetalhados,Numcrit, 7);
    SetLength(PrefDir,Numcrit);
    SetLength(MatrizConseq,numop, NUmcrit);
    SetLength(MatrizConseqNorm,numop,Numcrit);
    SetLength(Id_criterio,Numcrit);
    SetLength(Peso,Numcrit);
    SetLength(CritEspecial,Numcrit);
    Critsel := Numcrit-1;

    Nomecrit[Numcrit-1] := edtnomecrit.Text;
    niveis[numcrit-1] := strtoint(cmbbxniveiscriterios.Items[cmbbxniveiscriterios.ItemIndex]);
    prefdir[numcrit-1] := cmbbxprefdir.ItemIndex;
    critespecial[numcrit-1] := false;
    Ndetalhados[numcrit-1, 0] := edtn1.Text;
    Ndetalhados[numcrit-1, 1] := edtn2.Text;
    Ndetalhados[numcrit-1, 2] := edtn3.Text;
    Ndetalhados[numcrit-1, 3] := edtn4.Text;
    Ndetalhados[numcrit-1, 4] := edtn5.Text;
    Ndetalhados[numcrit-1, 5] := edtn6.Text;
    Ndetalhados[numcrit-1, 6] := edtn7.Text;
    With UserSession.zquery2 do
    begin


     edtn1.Clear;
     edtn2.Clear;
     edtn3.Clear;
     edtn4.Clear;
     edtn5.Clear;
     edtn6.Clear;
     edtn7.Clear;

     edtn1.Visible := False;
     edtn2.Visible := False;
     edtn3.Visible := False;
     edtn4.Visible := False;
     edtn5.Visible := False;
     edtn6.Visible := False;
     edtn7.Visible := False;

     cmbbxCriterios.ItemIndex := -1;
     cmbbxprefdir.ItemIndex := 1;
     cmbbxniveiscriterios.ItemIndex := -1;
     edtnomecrit.Clear;

     
     rgncriarnovocriterio.Visible := False;
     rgnvisualizarcriterios.Visible := True;
     lbldirecaopreferencia.Visible := False;
     lblnumerodeniveis.Visible := False;
     //WebApplication.ShowMessage('O critério foi cadastrado com sucesso!');
     NovoCriterio := true;
     //CriarQuestionarioCrit;
     rgncriterios.Visible := False;
    // rgnquestionariocrit.Visible := true;
    rgnquestoescrit.Visible := true;

    lblcritsel1.Text := Nomecrit[critsel];

    grdniveis.totalrows := niveis[critsel];
    grdniveis.RowCount := niveis[critsel];
    //grdniveis.ColumnCount := 2;



      grdniveis.ClearCells;
      for i := 0 to Niveis[critsel] - 1 do
      begin
      grdniveis.Cells[0,i] := inttostr(i);
      grdniveis.Cells[1,i] := Ndetalhados[critsel,i];
      end;



    grdquestionariocrit.TotalRows := NumOP;
    grdquestionariocrit.RowCount := NumOP;
    setlength(cmbbxconsequenciaop,numop);
    setlength(chlstconsequenciaop,numop);

    if critespecial[critsel] = true then
    begin
    grdniveis.columns[0].title := '<FONT face="Verdana" size="2"><B><P align="center">Categoria</P></B></FONT><FONT size="12"></FONT>';
    IWLabel55.caption := '*Selecione as categorias no qual cada operação se adequa em relação ao critério selecionado:';
    end
    else
    begin
    grdniveis.columns[0].title := '<FONT face="Verdana" size="2"><B><P align="center">Nível</P></B></FONT><FONT size="12"></FONT>';
    IWLabel55.caption := '*Selecione a opção no qual cada operação mais se adequa em relação ao critério selecionado:';
    end;
    for j := 0 to NumOP - 1 do
    begin
      grdquestionariocrit.Cells[0,j] := NomeOP[j];

      cmbbxconsequenciaop[j] := TiwComboBox.Create(Self);
      cmbbxconsequenciaop[j].parent := iwregion23;
      cmbbxconsequenciaop[j].width := grdquestionariocrit.Columns[1].Width+1;
      cmbbxconsequenciaop[j].Height := grdquestionariocrit.defaultrowheight;
      cmbbxconsequenciaop[j].top := 34+23*j;
      cmbbxconsequenciaop[j].left := 254;
      cmbbxconsequenciaop[j].bgColor := ClBlack;
      cmbbxconsequenciaop[j].font.Color := Clwhite;
      cmbbxconsequenciaop[j].zindex := 1000;


     for i := 0 to Niveis[critsel] - 1 do
    cmbbxconsequenciaop[j].Items.Add(inttostr(i));
    cmbbxconsequenciaop[j].ItemIndex := round(MatrizConseq[j,critsel])


    end;
      end;
   end;
end;
procedure TtfrmResultados.btneditarcritClick(Sender: TObject);
var
i, j : integer;
alfabeto : array of string;
begin
   With TIWUserSession(WebApplication.Data) do
   begin
    Nomecrit[critsel] := edtnomecrit.Text;
    niveis[critsel] := strtoint(cmbbxniveiscriterios.Items[cmbbxniveiscriterios.ItemIndex]);
    prefdir[critsel] := cmbbxprefdir.ItemIndex;
    Ndetalhados[critsel, 0] := edtn1.Text;
    Ndetalhados[critsel, 1] := edtn2.Text;
    Ndetalhados[critsel, 2] := edtn3.Text;
    Ndetalhados[critsel, 3] := edtn4.Text;
    Ndetalhados[critsel, 4] := edtn5.Text;
    Ndetalhados[critsel, 5] := edtn6.Text;
    Ndetalhados[critsel, 6] := edtn7.Text;
    for j := 0 to NumOP - 1 do
    begin
       if MatrizConseq[j,critsel] > Niveis[critsel] - 1 then
       matrizconseq[j,critsel] := niveis[critsel] - 1;
    end;
    With UserSession.zquery2 do
      begin
      SQL.Text := 'UPDATE criterios SET nome ="' + Nomecrit[critsel] + '" WHERE Id_criterio =' + Id_criterio[critsel];
      ExecSQL;
      SQL.Text := 'UPDATE criterios SET niveis ="' + inttostr(niveis[critsel]) + '" WHERE Id_criterio =' + Id_criterio[critsel];
      ExecSQL;
      if prefdir[critsel] = 1 then
      begin
      SQL.Text := 'UPDATE criterios SET prefdir ="' + 'maximização' + '" WHERE Id_criterio =' + Id_criterio[critsel];
      ExecSQL;
      end
      else
      begin
      SQL.Text := 'UPDATE criterios SET prefdir ="' + 'minimização' + '" WHERE Id_criterio =' + Id_criterio[critsel];
      ExecSQL;
      end;


      for i := 0 to NumOP - 1 do
      begin
       SQL.Text := 'UPDATE consequencias SET consequencia ="' + floattostr(matrizconseq[i,critsel]) + '" WHERE Id_criterio =' + Id_criterio[critsel] + ' AND Id_operacao = ' + Id_operacao[i];
       ExecSQL;
      end;



     SQL.Text := 'DELETE FROM niveis WHERE Id_criterio=' + Id_criterio[cmbbxcriterios.ItemIndex];
        ExecSQL;

    for i := 0 to Niveis[CritSel] - 1 do
     begin
       SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[CritSel] + '","' + Nomecrit[CritSel] + '","' + inttostr(i) + '", "' + Ndetalhados[CritSel,i] + '")';
      ExecSQL;
     end;

     With UserSession.zquery2 do
      begin
      SQL.Text := 'UPDATE usuario SET prob_finalizado ="' + 'Não' + '" WHERE Id_user =' + Id_user;
      ExecSQL;
      prob_finalizado := false;
      end;


       WebApplication.ShowMessage('O critério foi alterado com sucesso!');
      //rgncriarnovocriterio.visible := false;
      //rgnvisualizarcriterios.Visible := true;

      //cmbbxCriterios.ItemIndex := 0;
      //cmbbxCriteriosChange(self);


      NovoCriterio := False;
      //CriarQuestionarioCrit;

      rgncriterios.Visible := False;
    // rgnquestionariocrit.Visible := true;
    rgnquestoescrit.Visible := true;

    lblcritsel1.Text := Nomecrit[critsel];

    grdniveis.totalrows := niveis[critsel];
    grdniveis.RowCount := niveis[critsel];
    //grdniveis.ColumnCount := 2;
    setlength(alfabeto,7);
    alfabeto[0] := 'A';
    alfabeto[1] := 'B';
    alfabeto[2] := 'C';
    alfabeto[3] := 'D';
    alfabeto[4] := 'E';
    alfabeto[5] := 'F';
    alfabeto[6] := 'G';


      grdniveis.ClearCells;
      for i := 0 to Niveis[critsel] - 1 do
      begin
      grdniveis.Cells[0,i] := alfabeto[i];
      grdniveis.Cells[1,i] := Ndetalhados[critsel,i];
      end;



    grdquestionariocrit.TotalRows := NumOP;
    grdquestionariocrit.RowCount := NumOP;
    setlength(cmbbxconsequenciaop,numop);
    setlength(chlstconsequenciaop,numop);

 if critespecial[critsel] = true then
    begin
    grdniveis.columns[0].title := '<FONT face="Verdana" size="2"><B><P align="center">Categoria</P></B></FONT><FONT size="12"></FONT>';
    IWLabel55.caption := '*Selecione as categorias no qual cada operação se adequa em relação ao critério selecionado:';
    end
    else
    begin
    grdniveis.columns[0].title := '<FONT face="Verdana" size="2"><B><P align="center">Nível</P></B></FONT><FONT size="12"></FONT>';
    IWLabel55.caption := '*Selecione a opção no qual cada operação mais se adequa em relação ao critério selecionado:';
    end;

    for j := 0 to NumOP - 1 do
    begin
      grdquestionariocrit.Cells[0,j] := NomeOP[j];
      if critespecial[critsel] = false then
      begin
      cmbbxconsequenciaop[j] := TiwComboBox.Create(Self);
      cmbbxconsequenciaop[j].parent := iwregion23;
      cmbbxconsequenciaop[j].width := grdquestionariocrit.Columns[1].Width+1;
      cmbbxconsequenciaop[j].Height := grdquestionariocrit.defaultrowheight;
      cmbbxconsequenciaop[j].top := 34+23*j;
      cmbbxconsequenciaop[j].left := 254;
      cmbbxconsequenciaop[j].bgColor := ClBlack;
      cmbbxconsequenciaop[j].font.Color := Clwhite;
      cmbbxconsequenciaop[j].zindex := 1000;


     for i := 0 to Niveis[critsel] - 1 do
    cmbbxconsequenciaop[j].Items.Add(inttostr(i));
    cmbbxconsequenciaop[j].ItemIndex := round(MatrizConseq[j,critsel])
      end;
      if critespecial[critsel] = true then
      begin
       chlstconsequenciaop[j] := ttiwchecklist.Create(Self);
      chlstconsequenciaop[j].parent := iwregion23;
      chlstconsequenciaop[j].width := grdquestionariocrit.Columns[1].Width+1;
      chlstconsequenciaop[j].Height := grdquestionariocrit.defaultrowheight;
      chlstconsequenciaop[j].top := 36+23*j;
      chlstconsequenciaop[j].left := 254;
      chlstconsequenciaop[j].bgColor := clblack;
      chlstconsequenciaop[j].Color := ClBlack;
      chlstconsequenciaop[j].popupcolor := clwhite;
      chlstconsequenciaop[j].focuscolor := ClBlack;
      chlstconsequenciaop[j].font.Color := clwebgoldenrod;
      chlstconsequenciaop[j].font.Style := [fsBOLD];
      chlstconsequenciaop[j].zindex := 1000 - j;
      for i := 0 to Niveis[critsel] - 1 do
      chlstconsequenciaop[j].Items.Add(alfabeto[i]);
      end;
    end;
      end;
   end;
end;

procedure TtfrmResultados.btnEnterClick(Sender: TObject);
var
i,j, k,aux, Rlevels : integer;
DomMat : string;
pular : array of boolean;
auxniveis : array of real;
encontrado : boolean;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
  if rdbop.Checked = True then
  begin
    rgnopcoes.visible := false;
    rgnoperacoes.visible := True;
   // grdroperacoes.SelectColor := $f4d195;
    grdroperacoes.ClearRowSelect;
    grdroperacoes.Columns.Clear;

    while grdroperacoes.Columns.count <> 0  do
    grdroperacoes.Columns.Delete(0);

    grdroperacoes.Columns.Add;
    grdroperacoes.Columns[0].Title := '<FONT face="Verdana" size="2"><B><P align="center">Operação</P></B></FONT><FONT size="12"></FONT>';

    grdroperacoes.Columns.Add;
    grdroperacoes.Columns[1].Title := '<FONT face="Verdana" size="1"><B><P align="center">Data de Instauração do Inquerito</P></B></FONT><FONT size="12"></FONT>';

    grdroperacoes.Columns.Add;
    grdroperacoes.Columns[2].Title := '<FONT face="Verdana" size="1"><B><P align="center">Data do Fato</P></B></FONT><FONT size="12"></FONT>';

    for i := 0 to NumCrit - 1 do
    begin
    grdroperacoes.Columns.Add;
    grdroperacoes.Columns[i+3].Title := '<FONT  size="1"    face="Verdana"><B><P align="center">' + Nomecrit[i] + '</P></B></FONT>';
    end;

    grdroperacoes.TotalRows := NumOp;
    grdroperacoes.RowCount := NumOp;
    for j := 0 to Numop - 1 do
    begin
      grdroperacoes.Cells[0,j] := inttostr(j+1) + '- ' +  NomeOp[j];
      grdroperacoes.Cells[1,j] := DataOp[j];
      grdroperacoes.Cells[2,j] := DataFato[j];
    end;

    for i := 0 to NumCrit - 1 do
    for j := 0 to NumOp - 1 do
    begin
      grdroperacoes.Cells[i+3,j] := FloatToStr(MatrizConseq[j,i]);
    end;

    for i := 0 to grdroperacoes.Columns.count - 1 do
    grdroperacoes.Columns[i].ColumnHeaderClick := False;

    btnVisualizarOPClick(Self);
  end;
  if rdbcrit.Checked = True then
  begin
    rgnopcoes.visible := false;
    rgncriterios.visible := True;
    rgnvisualizarcriterios.Visible := true;
    rgncriarnovocriterio.Visible := False;

    lblnumerodeniveis.Visible := False;
    lbldirecaopreferencia.Visible := False;

    cmbbxCriterios.Clear;
    for i := 0 to NumCrit - 1 do
    cmbbxcriterios.Items.Add(NomeCrit[i]);
  end;

 if rdbpriori.Checked = true then
  begin

    if NumOP < 2 then
    begin
     WebApplication.ShowMessage('Você deve cadastrar mais de uma operação para prosseguir!');
     exit;
    end;

    if Numcrit < 2 then
    begin
     WebApplication.ShowMessage('Você deve cadastrar mais de um critério para prosseguir!');
     exit;
    end;

    grddetalhamentocrit.RowCount := Numcrit+1;
    grddetalhamentocrit.ColumnCount := 3;
    lstbxcriterios.Items.Clear;

    for i := 0 to Numcrit - 1 do
    lstbxcriterios.Items.Add(Nomecrit[i]);

    for j := 0 to grddetalhamentocrit.ColumnCount - 1 do
      begin
        grddetalhamentocrit.cell[0,j].Alignment := taCenter;
        grddetalhamentocrit.cell[0,j].wrap := True;
        grddetalhamentocrit.Cell[0,j].Font.Style := [fsbold];
        grddetalhamentocrit.Cell[0,j].BGCOLOR := clWebGOLDENROD;
        grddetalhamentocrit.Cell[0,j].fONT.Color := ClBlack;
      end;
      grddetalhamentocrit.Cell[0,0].Text := 'Critério';
      grddetalhamentocrit.Cell[0,1].Text := 'Pior Desempenho';
      grddetalhamentocrit.Cell[0,2].Text := 'Melhor Desempenho';


      lstbxordemcrit.Items.Clear;
      SetLength(MaxConseq, Numcrit);
      SetLength(minconseq, numcrit);
      for i := 1 to lstbxcriterios.Items.Count do
      begin
      MaxConseq[i-1] := -9999999999;
      MinConseq[i-1] := 99999999999;
      for j := 0 to NumOp - 1 do
      begin
      if matrizconseq[j,i-1] > maxconseq[i-1] then
      maxconseq[i-1] := matrizconseq[j,i-1];
      if matrizconseq[j,i-1] < minconseq[i-1] then
      minconseq[i-1] := matrizconseq[j,i-1];

      end;
      if maxconseq[i-1] = minconseq[i-1]  then
      maxconseq[i-1] := minconseq[i-1] + 0.000001;


      grddetalhamentocrit.Cell[i,0].Text := lstbxcriterios.Items[i-1];
      grddetalhamentocrit.Cell[i,1].Text := inttostr(round(minconseq[i-1]));
      grddetalhamentocrit.Cell[i,2].Text := inttostr(Round(maxconseq[i-1]));
      end;
      lstbxcriterioROC.Clear;
  lstbxrankROC.Clear;
  //rgnROC.Visible := false;
  With TIWUserSession(WebApplication.Data) do
  begin
  for i := 0 to numcrit - 1 do
  lstbxcriterioROC.Items.Add('C' + inttostr(i+1) + ': ' + Nomecrit[i]);


  end;
      if prob_finalizado = false then
      begin
      rgnopcoes.Visible := False;
      rgnelicitacao.Visible := true;
      Finalize(ordemcrit);
      gerar_grafico_overall;
      btnvisualizacaotabularClick(self);
      end
      else
      begin
       With TiwUserSession(WebApplication.Data) do
       begin

    Normaliza;
    finalize(resultadoop);
    //finalize(rankingfinal);
    finalize(matrizparapar);
    SetLength(resultadoop, numop);
    //SetLength(pular, numop);
    //SetLength(RankingFinal, NumOP);
    SetLength(MatrizParAPar, NumOP, NUmOP);
    for j := 0 to NumOp - 1 do
    begin
     // Pular[j] := false;
      for I := 0 to Numcrit - 1 do
      begin
        TiwUserSession(WebApplication.Data).ResultadoOP[j] := ResultadoOP[j] +
        TiwUserSession(WebApplication.Data).MatrizConseqNorm[j,i]*
        TiwUserSession(WebApplication.Data).Peso[i];
      end;
    end;

    NumNiveis := NumOP;


    for i := 0 to NumOP - 1 do
    begin

      for j := 0 to NUmOP - 1 do
      begin
        if i <> j then
        begin
          if ResultadoOP[i] > ResultadoOP[j] then
          begin
          TiwUserSession(WebApplication.Data).MatrizParaPar[i,j] := 1;
          end;
          if ResultadoOP[i] < ResultadoOP[j] then
          MatrizParapar[i,j] := 0;
          if ResultadoOP[i] = ResultadoOP[j] then
          begin
          MatrizParapar[i,j] := 2;
          end;

        end;
        DomMat := DomMat + inttostr(MatrizParapar[i,j]);
      end;
    end;

    auxNomeOP := '';
    for I := 0 to NumOP - 1 do
      begin
        auxNomeOP := auxNomeOP + '§' + NomeOP[i];
      end;

      setlength(auxniveis,1);
    auxniveis[0] := ResultadoOP[0];
    for I := 1 to NumOP - 1 do
    begin
    encontrado := false;
      for j := 0 to length(auxniveis) - 1 do
      begin
        if ResultadoOP[i] = auxniveis[j] then
        begin
        encontrado := true;
        break;
        end;
      end;
      if encontrado = false then
      begin
        setlength(auxniveis, length(auxniveis)+1);
        auxniveis[length(auxniveis)-1] := ResultadoOP[i];
      end;
    end;

    RLevels := length(auxniveis);
    // Teste unindo numAlt e Levels
      levels_numalt := IntToStr(RLevels) + '/' + IntToStr(NumOP);

      // Tirando o usuário da fila
      UserSession.ZQuery2.SQL.Text :=
        'DELETE FROM line_dh WHERE id_problem = "' + id_user + '"';
      ZQuery2.ExecSQL;

      UserSession.ZQuery2.SQL.Text :=
        'INSERT INTO line_dh (id_problem, id_user, dom_matrix, ah_matrix, NomeAlt, levels_numalt ) VALUES ("' + id_user + '", "' + ID_user + '", "' + DomMat + '", "' + '' + '", "' + auxNomeOP + '", "' + levels_numalt + '")';
      UserSession.ZQuery2.ExecSQL;

      UserSession.ZQuery2.SQL.Text :=
        'UPDATE desk SET dh_in_progress = "' + IntToStr(1) + '"';
      UserSession.ZQuery2.ExecSQL;

      tmrhasse.Enabled := true;


      // Indica que a aplicação desktop que constroi o diagrama deve ser executado
      // Informação lida pela aplicação desktop auxiliar - essa que vai chamar a outra aplicação que de fato constroi


    grdresultado.Clear;
    grdresultado.ColumnCount := 2;
    grdresultado.RowCount := NumOP+1;
    for j := 0 to grdresultado.ColumnCount - 1 do
      begin
        grdresultado.cell[0,j].Alignment := taCenter;
        grdresultado.cell[0,j].wrap := True;
        grdresultado.Cell[0,j].Font.Style := [fsbold];
        grdresultado.Cell[0,j].bgcolor := clWebGOLDENROD;
        grdresultado.Cell[0,j].Font.Color := clblack;
      end;
      grdresultado.Cell[0,0].text := 'Posição';
      grdresultado.Cell[0,1].text := 'Operação';


    k := 0;
    aux := 0;
    for i := 1 to NumOP do
    begin
      k := i - aux;
      for j := 0 to NumOP - 1 do
      begin
      if TiwUserSession(WebApplication.Data).RankingFinal[j] = i then
      begin
      grdresultado.Cell[k,0].Text := inttostr(k) + 'º';
      if grdresultado.Cell[k,1].Text = '' then
      grdresultado.Cell[k,1].Text := NomeOp[j]
      else
      begin
      grdresultado.Cell[k,1].Text :=  grdresultado.Cell[k,1].Text + ', ' + NomeOp[j];
      inc(aux);
      end;
      end;
      end;
    end;
    k := NumOP - aux;
    grdresultado.RowCount := k+1;


    txtordemcrit.Clear;
    SetLength(pular, numcrit);
    for i := 0 to Numcrit - 1 do
    pular[i] := false;


    for i := 0 to Numcrit - 1 do
    begin
    for j := 0 to Numcrit - 1 do
    begin
    if (ordemcrit[j] = i) and (pular[j] = false) then
    begin
    txtordemcrit.Lines.add(inttostr(i+1) + 'º - ' + Nomecrit[j]);
    pular[j] := true;
    //txtordemcrit.Lines.add('');
   end;
    end;
    end;
  end;
  rgnopcoes.Visible := False;
  rgnresultados.Visible := True;
  btnhasseClick(self);
      end;
  end;
  end;
end;

procedure TtfrmResultados.btnexcluircriterioClick(Sender: TObject);
var
i, j : integer;
begin
   With TIWUserSession(WebApplication.Data) do
   begin
   if cmbbxCriterios.ItemIndex = -1 then
    begin
     WebApplication.ShowMessage('Selecione um critério para continuar!');
     exit;
    end;
      With UserSession.zquery2 do
      begin
        SQL.Text := 'DELETE FROM criterios WHERE Id_criterio=' + Id_criterio[cmbbxcriterios.ItemIndex];
        ExecSQL;
        SQL.Text := 'DELETE FROM consequencias WHERE Id_criterio=' + Id_criterio[cmbbxcriterios.ItemIndex];
        ExecSQL;
        SQL.Text := 'DELETE FROM niveis WHERE Id_criterio=' + Id_criterio[cmbbxcriterios.ItemIndex];
        ExecSQL;
      end;

      With UserSession.zquery2 do
      begin
      SQL.Text := 'UPDATE usuario SET prob_finalizado ="' + 'Não' + '" WHERE Id_user =' + Id_user;
      ExecSQL;
      prob_finalizado := false;
      end;
    for i := cmbbxcriterios.ItemIndex to NumCrit - 2 do
    begin
      Nomecrit[i] := Nomecrit[i+1];
      Niveis[i] := Niveis[i+1];
      Prefdir[i] := Prefdir[i+1];
      Id_criterio[i] := Id_criterio[i+1];
      CritEspecial[i] := CritEspecial[i+1];
      for j := 0 to 6 do
      begin
        Ndetalhados[i,j] := Ndetalhados[i+1,j];
      end;
      for j := 0 to NumOp - 1 do
      begin
        MatrizConseq[j,i] := MatrizConseq[j,i+1];
      end;
    end;
    Numcrit := Numcrit - 1;
    SetLength(NomeCrit,Numcrit);
    SetLength(Niveis,Numcrit);
    SetLength(NDetalhados,Numcrit, 7);
    SetLength(PrefDir,Numcrit);
    SetLength(MatrizConseq,numop, NUmcrit);
    SetLength(MatrizConseqNorm,numop,Numcrit);
    SetLength(Id_criterio,Numcrit);
    SetLength(Peso,Numcrit);
    SetLength(CritEspecial,Numcrit);
    cmbbxcriterios.Items.Delete(cmbbxcriterios.ItemIndex);
    if numcrit > 0 then
    begin
    cmbbxcriterios.ItemIndex := 0;
    cmbbxCriteriosChange(self);
    end;
    WebApplication.ShowMessage('O critério foi excluído com sucesso!');
   end;
end;

procedure TtfrmResultados.btngorocClick(Sender: TObject);
var
i, j, k, aux : integer;
pular : array of boolean;
begin
  Smarter;
  With TiwUserSession(WebApplication.Data) do
  begin
    grdresultado.Clear;
    grdresultado.ColumnCount := 2;
    grdresultado.RowCount := NumOP+1;
    for j := 0 to grdresultado.ColumnCount - 1 do
      begin
        grdresultado.cell[0,j].Alignment := taCenter;
        grdresultado.cell[0,j].wrap := True;
        grdresultado.Cell[0,j].Font.Style := [fsbold];
        grdresultado.Cell[0,j].bgcolor := clWebGOLDENROD;
        grdresultado.Cell[0,j].Font.Color := clblack;
      end;
      grdresultado.Cell[0,0].text := 'Posição';
      grdresultado.Cell[0,1].text := 'Operação';


   k := 0;
    aux := 0;
    for i := 1 to NumOP do
    begin
      k := i - aux;
      for j := 0 to NumOP - 1 do
      begin
      if TiwUserSession(WebApplication.Data).RankingFinal[j] = i then
      begin
      grdresultado.Cell[k,0].Text := inttostr(k) + 'º';
      if grdresultado.Cell[k,1].Text = '' then
      grdresultado.Cell[k,1].Text := NomeOp[j]
      else
      begin
      grdresultado.Cell[k,1].Text :=  grdresultado.Cell[k,1].Text + ', ' + NomeOp[j];
      inc(aux);
      end;
      end;
      end;
    end;
    k := NumOP - aux;
    grdresultado.RowCount := k+1;
txtordemcrit.Clear;
    SetLength(pular, numcrit);
    //for i := 0 to Numcrit - 1 do
    //pular[i] := false;


    for i := 0 to Numcrit - 1 do
    begin
    //for j := 0 to Numcrit - 1 do
    //begin
    //if (ordemcrit[j] = i) and (pular[j] = false) then
    //begin
    txtordemcrit.Lines.add(inttostr(i+1) + 'º - ' + Nomecrit[ordemroc[i]]);
    //pular[j] := true;
    //txtordemcrit.Lines.add('');
   //end;
    //end;
    end;
  end;
  rgnelicitacao.Visible := False;
  rgnresultados.Visible := True;
  btnvisualizacaotabular.Visible := True;
  btnvisualizacaografica.Visible := True;
  btnhasseClick(self);
end;

procedure TtfrmResultados.btnhasseClick(Sender: TObject);
begin
 rgnresultadotabular.visible := false;
  rgnresultadohasse.visible := true;
  if rgnresultadohasse.visible = true then
  begin
    btnhasse.BackColor := clwebsilver;
    btnhasse.backcolorto := clwebgray;
    btntabela.BackColor := clWebGoldenRod;
    btntabela.backcolorto := clwebPeachPuff;
    rgnbiresultado.height := rgnresultadohasse.Height + 100;
  end
  else
  begin
    btnhasse.BackColor := clWebGoldenRod;
    btnhasse.backcolorto := clwebPeachPuff;
    btntabela.BackColor := clwebsilver;
    btntabela.backcolorto := clwebgray;
  end;
end;


procedure TtfrmResultados.btnProsseguirClick(Sender: TObject);
var
i : integer;
begin
With TIWUserSession(WebApplication.Data) do
  begin
    if edtnomeop.Text = '' then
    begin
    WebApplication.ShowMessage('Preencha corretamente as informações!');
    Exit;
    end;
    for i := 0 to NumOP - 1 do
    begin
      if edtnomeop.Text = nomeop[i] then
      begin
        WebApplication.ShowMessage('Já existe uma operação com este nome. Preencha novamente o campo!');
        Exit;
      end;
    end;
    Numop := Numop + 1;
    SetLength(NomeOp,numop);
    SetLength(DataOp,numop);
    SetLength(DetalhesOP,numop);
    SetLength(RankingFinal,numop);
    SetLength(MatrizConseq,numop, NUmcrit);
    SetLength(MatrizConseqNorm,numop,Numcrit);
    SetLength(Id_Operacao,numop);
    SetLength(DataFato, NUmOP);

    NomeOp[NumOp-1] := edtnomeop.Text;
    DataOP[NumOp-1] := edtdia.Text + '/' + edtmes.Text + '/' + edtano.Text;
    DataFato[NumOp-1] := edtdiafato.Text + '/' + edtmesfato.Text + '/' + edtanofato.Text;
    detalhesop[NumOP-1] := mmDetalhesOP.Text;
    lbloperacaosel1.Caption := edtnomeop.Text;

    if (NomeOp[NumOp-1] = '') or (DataOP[NumOp-1] = '') or (DataFato[NumOp-1] = '') or (detalhesop[NumOP-1] = '') then
    begin
      NumOP := Numop-1;
      WebApplication.ShowMessage('Preencha corretamente as informações!');
      exit;
    end;

    With UserSession.zquery2 do
    begin
      SQL.Text :=
      'INSERT INTO operacoes (Id_user, nome, data, data_fato, detalhamento_operacao) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + edtNomeOP.Text + '","' + edtdia.Text + '/' + edtmes.Text + '/' + edtano.Text + '","' + edtdiafato.Text + '/' + edtmesfato.Text + '/' + edtanofato.Text + '", "' + mmDetalhesOP.Text + '")';
      ExecSQL;

      SQL.Text := 'SELECT Id FROM operacoes WHERE Id_user="' + ID_user + '" AND nome = "' + edtnomeop.Text + '"';
        ExecSQL;
        open;
        Id_operacao[NumOP-1] := FieldByName('Id').AsString;
        close;

      for I := 0 to numcrit - 1 do
      begin
       SQL.Text :=
      'INSERT INTO consequencias (Id_user, Id_operacao, Id_criterio) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_operacao[NumOP-1] + '","' + Id_criterio[i] + '")';
      ExecSQL;
      end;
    end;

    With UserSession.zquery2 do
      begin
      SQL.Text := 'UPDATE usuario SET prob_finalizado ="' + 'Não' + '" WHERE Id_user =' + Id_user;
      ExecSQL;
      prob_finalizado := false;
      end;

    grdroperacoes.RowCount := NumOP;
    grdroperacoes.TotalRows := NumOP;

    grdroperacoes.Cells[0,NumOP-1] := IntToStr(NumOP) + '- ' + edtnomeop.Text;
    grdroperacoes.Cells[1,NumOP-1] := edtdia.Text + '/' + edtmes.Text + '/' + edtano.Text;
    grdroperacoes.Cells[2,NumOP-1] := edtdiafato.Text + '/' + edtmesfato.Text + '/' + edtanofato.Text;
    for I := 0 to Numcrit - 1 do
    grdroperacoes.cells[3+i,NumOp-1] := FloatToStr(MatrizConseq[NumOp-1,i]);


    edtnomeop.Clear;
    //edtdataop.Clear;
    edtdia.Clear;
    edtmes.Clear;
    edtano.Clear;
    edtdiafato.Clear;
    edtmesfato.Clear;
    edtanofato.Clear;
    mmDetalhesOP.Clear;
    OpSel := NumOp -1;



  end;
rgnoperacoes.Visible := false;
rgnquestionario.Visible := True;
//lbloperacaosel1.Caption := edtnomeop.Text;
//txtpergunta.Lines[0] := 'Selecione a opção que a operação "' + lbloperacaosel1.Caption + '" mais se adequa em relação ao critério "Dano à Pessoa Humana".';
criarquestionario;
end;



procedure TtfrmResultados.btnreiniciarClick(Sender: TObject);
var
i,j, aux : integer;
begin
    With TiwUserSession(WebApplication.Data) do
    begin
    grddetalhamentocrit.RowCount := Numcrit+1;
    grddetalhamentocrit.ColumnCount := 3;
    lstbxcriterios.Items.Clear;
    Finalize(ordemcrit);
    finalize(auxnomecrit);
    setlength(auxnomecrit, numcrit);
    for i := 0 to Numcrit - 1 do
    begin
    auxnomecrit[i] := nomecrit[i];
    lstbxcriterios.Items.Add('C'+ inttostr(I+1) + ': '+ Nomecrit[i]);
    end;
    for j := 0 to grddetalhamentocrit.ColumnCount - 1 do
      begin
        grddetalhamentocrit.cell[0,j].Alignment := taCenter;
        grddetalhamentocrit.cell[0,j].wrap := True;
        grddetalhamentocrit.Cell[0,j].Font.Style := [fsbold];
        grddetalhamentocrit.Cell[0,j].BGCOLOR := clWebGOLDENROD;
        grddetalhamentocrit.Cell[0,j].fONT.Color := ClBlack;
      end;
      grddetalhamentocrit.Cell[0,0].Text := 'Critério';
      //grddetalhamentocrit.Cell[0,1].Text := 'Direção de Preferência';
      grddetalhamentocrit.Cell[0,1].Text := 'Menor Criticidade';
      grddetalhamentocrit.Cell[0,2].Text := 'Maior Criticidade';


      lstbxordemcrit.Items.Clear;
      SetLength(MaxConseq, Numcrit);
      SetLength(minconseq, numcrit);
      for i := 1 to lstbxcriterios.Items.Count do
      begin
      MaxConseq[i-1] := -9999999999;
      MinConseq[i-1] := 99999999999;
      for j := 0 to NumOp - 1 do
      begin
      if matrizconseq[j,i-1] > maxconseq[i-1] then
      maxconseq[i-1] := matrizconseq[j,i-1];
      if matrizconseq[j,i-1] < minconseq[i-1] then
      minconseq[i-1] := matrizconseq[j,i-1];

      end;
      if maxconseq[i-1] = minconseq[i-1]  then
      maxconseq[i-1] := minconseq[i-1] + 0.000001;


      grddetalhamentocrit.Cell[i,0].Text := lstbxcriterios.Items[i-1];
      grddetalhamentocrit.Cell[i,1].Text := inttostr(round(minconseq[i-1]));
      grddetalhamentocrit.Cell[i,2].Text := inttostr(Round(maxconseq[i-1]));
      end;
     btnverresultados.enabled := false;
     lstbxcriteriosChange(self);
  end;



end;

procedure TtfrmResultados.BtnrestartClick(Sender: TObject);
var
i : integer;
begin
  lstbxcriterioROC.Clear;
  lstbxrankROC.Clear;
  //rgnROC.Visible := false;
  With TIWUserSession(WebApplication.Data) do
  begin
  for i := 0 to numcrit - 1 do
  begin
  lstbxcriterioROC.Items.Add('C' + inttostr(i+1) + ': ' + Nomecrit[i]);
  vetorgrafico[i] := 0;
  end;



  Finalize(ordemcrit);
  gerar_grafico_overall;
  btngoroc.Enabled := False;

  end;
end;

procedure TtfrmResultados.btnsalvarClick(Sender: TObject);
var
i, j, k, aux2 : integer;
num, aux : double;
preencherordem : array of double;
auxordemcrit : array of integer;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    Finalize(pesosavdireta);
    setlength(pesosavdireta, numcrit);
    setlength(ordemcrit, numcrit);
    for i := 0 to Numcrit - 1 do
    begin
      edtcritavdireta[i].Text := TrocaVP(edtcritavdireta[i].Text);
      if TryStrToFloat(edtcritavdireta[i].Text, num) = true then
      pesosavdireta[i] := StrToFloat(edtcritavdireta[i].Text)
      else
      begin
      WebApplication.ShowMessage('Preencha corretamente as informações!');
      exit;
      end;
    end;
    for i := 0 to Numcrit - 1 do
    begin
      edtcritavdireta[i].Enabled := False;
      aux := aux + pesosavdireta[i];
    end;

    for i := 0 to numcrit - 1 do
    begin
    pesosavdireta[i] := pesosavdireta[i]/aux;
    edtcritavdireta[i].Text := FormatFloat('0.00', pesosavdireta[i]);
    end;


    setlength(preencherordem,1);
    preencherordem[0] := pesosavdireta[0];
    for i := 1 to Numcrit - 1 do
    begin
      if pesosavdireta[i] <> preencherordem[length(preencherordem)-1] then
      begin
        setlength(preencherordem,length(preencherordem)+1);
        //preencherordem[length(preencherordem)-1] := preencherordem[length(preencherordem)-2];
        preencherordem[length(preencherordem)-1] := pesosavdireta[i];
      end;
    end;
    SetLength(auxordemcrit, length(preencherordem));
    for i := 0 to length(preencherordem) - 1 do
    begin
      auxordemcrit[i] := 0;
      for j := 0 to length(preencherordem) - 1 do
      begin
        if i <> j then
        begin
          if preencherordem[i] < preencherordem[j] then
          inc(auxordemcrit[i]);
        end;
      end;

    end;


    for i := 0 to Numcrit - 1 do
    begin
    for j := 0 to length(preencherordem) - 1 do
    begin
      if pesosavdireta[i] = preencherordem[j] then
      TIWUserSession(WebApplication.Data).ordemcrit[i] := auxordemcrit[j];
    end;
    grdinputmanual.Cell[2,i+1].Text := inttostr(ordemcrit[i]+1);
     end;

     for i := 0 to TIWUserSession(WebApplication.Data).Numcrit - 1 do
  begin
   With UserSession.zquery2 do
      begin
      SQL.Text := 'UPDATE criterios SET peso ="' + floattostr(TIWUserSession(WebApplication.Data).pesosavdireta[i]) + '" WHERE Id_criterio=' + TIWUserSession(WebApplication.Data).Id_criterio[i];
      ExecSQL;
      end;
  end;




    btnverresultadosavdireta.Enabled := True;
    btnsalvar.Enabled := False;
  end;
end;

procedure TtfrmResultados.btnsalvaropClick(Sender: TObject);
var
i, j : integer;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    Numop := Numop + 1;
    SetLength(NomeOp,numop);
    SetLength(DataOp,numop);
    SetLength(datafato,numop);
    SetLength(DetalhesOP,numop);
    SetLength(RankingFinal,numop);
    SetLength(MatrizConseq,numop, NUmcrit);
    SetLength(MatrizConseqNorm,numop,Numcrit);
    SetLength(Id_Operacao,numop);

    NomeOp[NumOp-1] := edtnomeop.Text;
    DataOP[NumOp-1] := edtdia.Text + '/' + edtmes.Text + '/' + edtano.Text;
    DataFato[NumOp-1] := edtdiafato.Text + '/' + edtmesfato.Text + '/' + edtanofato.Text;
    detalhesop[NumOP-1] := mmDetalhesOP.Text;

    With UserSession.zquery2 do
    begin
      SQL.Text :=
      'INSERT INTO operacoes (Id_user, nome, data, detalhamento_operacao) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + edtNomeOP.Text + '","' + edtdia.Text + '/' + edtmes.Text + '/' + edtano.Text + '", "' + mmDetalhesOP.Text + '")';
      ExecSQL;

      SQL.Text := 'SELECT Id FROM operacoes WHERE Id_user="' + ID_user + '" AND nome = "' + edtnomeop.Text + '"';
        ExecSQL;
        open;
        Id_operacao[NumOP-1] := FieldByName('Id').AsString;
        close;

      for I := 0 to numcrit - 1 do
      begin
       SQL.Text :=
      'INSERT INTO consequencias (Id_user, Id_operacao, Id_criterio) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_operacao[NumOP-1] + '","' + Id_criterio[i] + '")';
      ExecSQL;
      end;
    end;

    With UserSession.zquery2 do
      begin
      SQL.Text := 'UPDATE usuario SET prob_finalizado ="' + 'Não' + '" WHERE Id_user =' + Id_user;
      ExecSQL;
      prob_finalizado := false;
      end;

    grdroperacoes.RowCount := NumOP;
    grdroperacoes.TotalRows := NumOP;

    grdroperacoes.Cells[0,NumOP-1] := edtnomeop.Text;
    grdroperacoes.Cells[1,NumOP-1] := edtdia.Text + '/' + edtmes.Text + '/' + edtano.Text;
    for I := 0 to Numcrit - 1 do
    grdroperacoes.cells[2+i,NumOp-1] := FloatToStr(MatrizConseq[NumOp-1,i]);


    edtnomeop.Clear;
    edtdia.Clear;
    edtmes.Clear;
    edtano.Clear;
    mmDetalhesOP.Clear;
    WebApplication.ShowMessage('A operação foi cadastrada com sucesso!');
    btnvisualizarOPClick(self);



  end;

end;

procedure TtfrmResultados.btnsalvarquestionarioClick(Sender: TObject);
var
i,j,k : integer;
begin
  With TiwUserSession(WebApplication.Data) do
  begin
    With UserSession.zquery2 do
    begin
      for i := 0 to Numcrit - 1 do
      begin
      if Nomecrit[i] <> 'Repercussão e Extensão do Crime' then
      begin
      SQL.Text := 'UPDATE consequencias SET consequencia ="' + inttostr(rdgniveis[i].itemindex) + '" WHERE Id_operacao=' + Id_Operacao[OpSel] +' AND Id_criterio =' + Id_criterio[i];
      ExecSQL;
      matrizconseq[opsel,i] := rdgniveis[i].itemindex;
      end
      else
      begin
        k := 0;
        for j := 0 to niveis[i] - 1 do
        begin
          if clsbtxCriterioEspecial.Selected[j] = true then
          inc(k);
        end;
        SQL.Text := 'UPDATE consequencias SET consequencia ="' + inttostr(k) + '" WHERE Id_operacao=' + Id_Operacao[OpSel] +' AND Id_criterio =' + Id_criterio[i];
      ExecSQL;
       matrizconseq[opsel,i] := k;
      end;
      end;
    end;

    rgnquestionario.visible := false;
    rgnoperacoes.visible := True;
   // grdroperacoes.SelectColor := $f4d195;
    grdroperacoes.ClearRowSelect;
    grdroperacoes.Columns.Clear;

    while grdroperacoes.Columns.count <> 0  do
    grdroperacoes.Columns.Delete(0);

    grdroperacoes.Columns.Add;
    grdroperacoes.Columns[0].Title := '<FONT face="Verdana" size="2"><B><P align="center">Operação</P></B></FONT><FONT size="12"></FONT>';

    grdroperacoes.Columns.Add;
    grdroperacoes.Columns[1].Title := '<FONT face="Verdana" size="1"><B><P align="center">Data de Instauração do Inquerito</P></B></FONT><FONT size="12"></FONT>';

    grdroperacoes.Columns.Add;
    grdroperacoes.Columns[2].Title := '<FONT face="Verdana" size="1"><B><P align="center">Data do Fato</P></B></FONT><FONT size="12"></FONT>';

    for i := 0 to NumCrit - 1 do
    begin
    grdroperacoes.Columns.Add;
    grdroperacoes.Columns[i+3].Title := '<FONT  size="1"    face="Verdana"><B><P align="center">' + Nomecrit[i] + '</P></B></FONT>';
    end;

    grdroperacoes.TotalRows := NumOp;
    grdroperacoes.RowCount := NumOp;
    for j := 0 to Numop - 1 do
    begin
      grdroperacoes.Cells[0,j] := inttostr(j+1) + '- ' +  NomeOp[j];
      grdroperacoes.Cells[1,j] := DataOp[j];
      grdroperacoes.Cells[2,j] := DataFato[j];
    end;

    for i := 0 to NumCrit - 1 do
    for j := 0 to NumOp - 1 do
    begin
      grdroperacoes.Cells[i+3,j] := FloatToStr(MatrizConseq[j,i]);
    end;

    for i := 0 to grdroperacoes.Columns.count - 1 do
    grdroperacoes.Columns[i].ColumnHeaderClick := False;

    deletarquestionario;
    btnVisualizarOPClick(Self);
  end;

end;




procedure TtfrmResultados.btnsalvarquestoescritClick(Sender: TObject);
var
i,j,k : integer;
begin
  With TiwUserSession(WebApplication.Data) do
  begin
    With UserSession.zquery2 do
    begin
      if NovoCriterio = True then
      begin
       if PrefDir[critsel] = 1 then
    begin
     SQL.Text :=
      'INSERT INTO criterios (Id_user, nome, prefdir, niveis) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Nomecrit[critsel] + '","' + 'maximização' + '", "' + IntToStr(Niveis[Numcrit-1]) + '")';
    ExecSQL;
    end
    else
    begin
     SQL.Text :=
      'INSERT INTO criterios (Id_user, nome, prefdir, niveis) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Nomecrit[critsel] + '","' + 'minimização' + '", "' + IntToStr(Niveis[critsel]) + '")';
    ExecSQL;
    end;
    SQL.Text := 'SELECT Id_criterio FROM criterios WHERE Id_user="' + ID_user + '" AND nome = "' + Nomecrit[critsel] + '"';
        ExecSQL;
        open;
        Id_Criterio[critsel] := FieldByName('Id_criterio').AsString;
        close;

      for I := 0 to numop - 1 do
      begin
       SQL.Text :=
      'INSERT INTO consequencias (Id_user, Id_operacao, Id_criterio) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_operacao[i] + '","' + Id_criterio[critsel] + '")';
      ExecSQL;
      end;


     for i := 0 to Niveis[Numcrit-1] - 1 do
     begin
       SQL.Text :=
      'INSERT INTO niveis (Id_user, Id_criterio, nome_criterio, nivel, detalhamento) VALUES ("' + TIWUserSession(WebApplication.Data).ID_user + '","' + Id_criterio[critsel] + '","' + Nomecrit[critsel] + '","' + inttostr(i) + '", "' + Ndetalhados[critsel,i] + '")';
      ExecSQL;
     end;


      SQL.Text := 'UPDATE usuario SET prob_finalizado ="' + 'Não' + '" WHERE Id_user =' + Id_user;
      ExecSQL;
      prob_finalizado := false;

      end;

      if critespecial[critsel] = false then
      begin
      for i := 0 to NumOP - 1 do
      begin
      SQL.Text := 'UPDATE consequencias SET consequencia ="' + inttostr(cmbbxconsequenciaop[i].itemindex) + '" WHERE Id_operacao=' + Id_Operacao[i] +' AND Id_criterio =' + Id_criterio[critsel];
      ExecSQL;
      matrizconseq[i,critsel] := cmbbxconsequenciaop[i].itemindex;
      end;
      end
      else
      begin
       for i := 0 to NumOP - 1 do
      begin
      k := 0;
      for j := 0 to Niveis[critsel] - 1 do
      begin
      if chlstconsequenciaop[i].Selected[j] = true then
      inc(k);
      end;
      SQL.Text := 'UPDATE consequencias SET consequencia ="' + inttostr(k) + '" WHERE Id_operacao=' + Id_Operacao[i] +' AND Id_criterio =' + Id_criterio[critsel];
      ExecSQL;
      matrizconseq[i,critsel] := k;
      end;
      end;
    end;

    for i := 0 to NumOP - 1 do
    begin
      if critespecial[critsel] = true then
      chlstconsequenciaop[i].destroy;
      if critespecial[critsel] = false then
      cmbbxconsequenciaop[i].destroy;
    end;


    cmbbxcriterios.Clear;
     for i := 0 to Numcrit - 1 do
     cmbbxcriterios.Items.Add(NomeCrit[i]);

    rgnquestoescrit.visible := false;
    rgncriterios.visible := True;



    //deletarquestionariocrit;
    //btnVisualizarOPClick(Self);
  end;

end;

procedure TtfrmResultados.rgncriteriosCreate(Sender: TObject);
begin
  rgncriterios.Align := AlClient;
end;

procedure TtfrmResultados.rgnelicitacaoCreate(Sender: TObject);
begin
rgnelicitacao.Align := AlClient;
end;

procedure TtfrmResultados.rgninputmanualCreate(Sender: TObject);
begin
rgninputmanual.Align := AlClient;
end;

procedure TtfrmResultados.rgnopcoesCreate(Sender: TObject);
begin
  rgnopcoes.Align := AlClient;
end;

procedure TtfrmResultados.rgnoperacoesCreate(Sender: TObject);
begin
  rgnoperacoes.Align := AlClient;
end;

procedure TtfrmResultados.rgnquestionarioCreate(Sender: TObject);
begin
rgnquestionario.Align := Alclient;
end;



procedure TtfrmResultados.rgnquestoescritCreate(Sender: TObject);
begin
rgnquestoescrit.Align := AlClient;
end;

procedure TtfrmResultados.rgnresultadosCreate(Sender: TObject);
begin
rgnresultados.Align := AlClient;
end;


procedure TtfrmResultados.btnVisualizarOPClick(Sender: TObject);
begin
  rgnvisualizarOP.visible := True;
  rgncriarOP.visible := false;
  if rgnvisualizarOP.visible = True then
  begin
    btnVisualizarOP.BackColor := clwebsilver;
    btnVisualizarOP.backcolorto := clwebgray;
    btnCadastrarOP.BackColor := clWebGoldenRod;
    btnCadastrarOP.backcolorto := clwebPeachPuff;
  end
  else
  begin
    btnVisualizarOP.BackColor := clWebGoldenRod;
    btnVisualizarOP.backcolorto := clwebPeachPuff;
    btnCadastrarOP.BackColor := clwebsilver;
    btnCadastrarOP.backcolorto := clwebgray;
  end;
  

end;

procedure TtfrmResultados.cmbbxCriteriosChange(Sender: TObject);
var
  i, j: integer;
  alfabeto : array of string;
begin
  With TIWUserSession(WebApplication.Data) do
  begin

  grdcriterios.RowCount := Niveis[cmbbxcriterios.ItemIndex];
  grdcriterios.TotalRows := Niveis[cmbbxcriterios.ItemIndex];

  setlength(alfabeto,7);
    alfabeto[0] := 'A';
    alfabeto[1] := 'B';
    alfabeto[2] := 'C';
    alfabeto[3] := 'D';
    alfabeto[4] := 'E';
    alfabeto[5] := 'F';
    alfabeto[6] := 'G';

  if critespecial[cmbbxcriterios.ItemIndex] = true then
    begin
    grdcriterios.columns[0].title := '<FONT face="Verdana" size="2"><B><P align="center">Categoria</P></B></FONT><FONT size="12"></FONT>';
   // IWLabel55.caption := '*Selecione as categorias no qual cada operação se adequa em relação ao critério selecionado:';
    end
    else
    begin
    grdcriterios.columns[0].title := '<FONT face="Verdana" size="2"><B><P align="center">Níveis</P></B></FONT><FONT size="12"></FONT>';
   // IWLabel55.caption := '*Selecione a opção no qual cada operação mais se adequa em relação ao critério selecionado:';
    end;

  if critespecial[cmbbxcriterios.ItemIndex] = false then
  begin
  for i := 0 to Niveis[cmbbxcriterios.ItemIndex] - 1 do
  begin
    grdcriterios.Cells[0, i] := 'Nível ' + inttostr(i);
    grdcriterios.Cells[1, i] := Ndetalhados[cmbbxcriterios.ItemIndex, i];
  end;
  end
  else
  begin
   for i := 0 to Niveis[cmbbxcriterios.ItemIndex] - 1 do
  begin
    grdcriterios.Cells[0, i] := alfabeto[i];
    grdcriterios.Cells[1, i] := Ndetalhados[cmbbxcriterios.ItemIndex, i];
  end;
  end;

   lblnumerodeniveis.Visible := True;
   lbldirecaopreferencia.Visible := True;

   if prefdir[cmbbxcriterios.ItemIndex] = 1 then
   lbldirecaopreferencia.Caption := 'Maximização'
   else
   lbldirecaopreferencia.Caption := 'Minimização';

   lblnumerodeniveis.Caption := IntToStr(Niveis[cmbbxcriterios.ItemIndex]);




  end;

end;

procedure TtfrmResultados.cmbbxniveiscriteriosChange(Sender: TObject);
var
i,j : integer;
alfabeto : array of string;
begin
  rgnpreencherniveis.Visible := True;
  if cmbbxniveiscriterios.itemindex = 0 then
  begin
  grdpreencherniveis.RowCount := 2;
  grdpreencherniveis.TotalRows := 2;
  edtn1.Visible := True;
  edtn2.Visible := True;
  edtn3.Visible := False;
  edtn4.Visible := False;
  edtn5.Visible := False;
  edtn6.Visible := False;
  edtn7.Visible := False;

  edtn1.Clear;
  edtn2.Clear;
  //edtn3.Clear;
  //edtn4.Clear;
  //edtn5.Clear;
  //edtn6.Clear;
  //edtn7.Clear;
  end;

  if cmbbxniveiscriterios.itemindex = 1 then
  begin
  grdpreencherniveis.RowCount := 3;
  grdpreencherniveis.TotalRows := 3;
  edtn1.Visible := True;
  edtn2.Visible := True;
  edtn3.Visible := True;
  edtn4.Visible := False;
  edtn5.Visible := False;
  edtn6.Visible := False;
  edtn7.Visible := False;

  edtn1.Clear;
  edtn2.Clear;
  edtn3.Clear;
  //edtn4.Clear;
  //edtn5.Clear;
  //edtn6.Clear;
  //edtn7.Clear;
  end;

  if cmbbxniveiscriterios.itemindex = 2 then
  begin
  grdpreencherniveis.RowCount := 4;
  grdpreencherniveis.TotalRows := 4;
  edtn1.Visible := True;
  edtn2.Visible := True;
  edtn3.Visible := True;
  edtn4.Visible := True;
  edtn5.Visible := False;
  edtn6.Visible := False;
  edtn7.Visible := False;

  edtn1.Clear;
  edtn2.Clear;
  edtn3.Clear;
  edtn4.Clear;
  //edtn5.Clear;
  //edtn6.Clear;
  //edtn7.Clear;
  end;

  if cmbbxniveiscriterios.itemindex = 3 then
  begin
  grdpreencherniveis.RowCount := 5;
  grdpreencherniveis.TotalRows := 5;
  edtn1.Visible := True;
  edtn2.Visible := True;
  edtn3.Visible := True;
  edtn4.Visible := True;
  edtn5.Visible := True;
  edtn6.Visible := False;
  edtn7.Visible := False;

  edtn1.Clear;
  edtn2.Clear;
  edtn3.Clear;
  edtn4.Clear;
  edtn5.Clear;
  //edtn6.Clear;
  //edtn7.Clear;
  end;

  if cmbbxniveiscriterios.itemindex = 4 then
  begin
  grdpreencherniveis.RowCount := 6;
  grdpreencherniveis.TotalRows := 6;
  edtn1.Visible := True;
  edtn2.Visible := True;
  edtn3.Visible := True;
  edtn4.Visible := True;
  edtn5.Visible := True;
  edtn6.Visible := True;
  edtn7.Visible := False;

  edtn1.Clear;
  edtn2.Clear;
  edtn3.Clear;
  edtn4.Clear;
  edtn5.Clear;
  edtn6.Clear;
  //edtn7.Clear;
  end;

  if cmbbxniveiscriterios.itemindex = 5 then
  begin
  grdpreencherniveis.RowCount := 7;
  grdpreencherniveis.TotalRows := 7;
  edtn1.Visible := True;
  edtn2.Visible := True;
  edtn3.Visible := True;
  edtn4.Visible := True;
  edtn5.Visible := True;
  edtn6.Visible := True;
  edtn7.Visible := True;
  end;

  setlength(alfabeto,7);
    alfabeto[0] := 'A';
    alfabeto[1] := 'B';
    alfabeto[2] := 'C';
    alfabeto[3] := 'D';
    alfabeto[4] := 'E';
    alfabeto[5] := 'F';
    alfabeto[6] := 'G';
  if (TIWUserSession(WebApplication.Data).CritEspecial[TIWUserSession(WebApplication.Data).critsel] = false) or (TIWUserSession(WebApplication.Data).NovoCriterio = true) then
  begin
  for i := 0 to grdpreencherniveis.RowCount - 1 do
  begin
    grdpreencherniveis.Cells[0, i] := 'Nível ' + inttostr(i);
  end;
  end
  else
  begin
   for i := 0 to grdpreencherniveis.RowCount - 1 do
  begin
    grdpreencherniveis.Cells[0, i] := alfabeto[i];
  end;
  end;

end;



procedure TtfrmResultados.IWAppFormCreate(Sender: TObject);
var
i,j,k : integer;
preencherordem : array of double;
auxordemcrit : array of integer;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
      With UserSession.zquery2 do
      begin
        SQL.Text := 'SELECT prob_finalizado FROM usuario WHERE Id_user ="' + Id_User + '"';
        ExecSQL;
        open;
        if FieldByName('prob_finalizado').Asstring = 'Sim' then
        prob_finalizado := true
        else
        prob_finalizado := False;
        Close;



        SQL.Text := 'SELECT count(*) FROM criterios WHERE Id_user ="' + Id_User + '"';
        ExecSQL;
        open;
        Numcrit := FieldByName('count(*)').AsInteger;
        SetLength(TIWUserSession(WebApplication.Data).Niveis, Numcrit);
        SetLength(TIWUserSession(WebApplication.Data).Nomecrit, Numcrit);
        SetLength(TIWUserSession(WebApplication.Data).Peso, Numcrit);
        SetLength(TIWUserSession(WebApplication.Data).Id_Criterio, Numcrit);
        SetLength(TIWUserSession(WebApplication.Data).prefdir, Numcrit);
        SetLength(TIWUserSession(WebApplication.Data).ordemcrit, Numcrit);
        SetLength(TIWUserSession(WebApplication.Data).CritEspecial, Numcrit);
        close;
        SQL.Text := 'SELECT * FROM criterios WHERE Id_user ="' + Id_User + '"';
        ExecSQL;
        open;
        First;
        i := 0;
        while not eof do
        begin
          Id_Criterio[i] := IntToStr(FieldByName('Id_criterio').AsInteger);
          NomeCrit[i] := FieldByName('nome').Asstring;
          if FieldByName('prefdir').Asstring = 'maximização' then
          prefdir[i] := 1
          else
          prefdir[i] := 0;
          if FieldByName('somatorio').Asstring = 'Sim' then
          CritEspecial[i] := true
          else
          CritEspecial[i] := false;
          Niveis[i] := StrToInt(FieldByName('niveis').Asstring);
          if prob_finalizado  = true  then
          peso[i] := StrToFloat(FieldByName('peso').Asstring);

          inc(i);
          Next;
        end;
        Close;

    if prob_finalizado = true then
    begin
    setlength(preencherordem,1);
    preencherordem[0] := peso[0];
    for i := 1 to Numcrit - 1 do
    begin
      if peso[i] <> preencherordem[length(preencherordem)-1] then
      begin
        setlength(preencherordem,length(preencherordem)+1);
        //preencherordem[length(preencherordem)-1] := preencherordem[length(preencherordem)-2];
        preencherordem[length(preencherordem)-1] := peso[i];
      end;
    end;
    SetLength(auxordemcrit, length(preencherordem));
    for i := 0 to length(preencherordem) - 1 do
    begin
      auxordemcrit[i] := 0;
      for j := 0 to length(preencherordem) - 1 do
      begin
        if i <> j then
        begin
          if preencherordem[i] < preencherordem[j] then
          inc(auxordemcrit[i]);
        end;
      end;

    end;


    for i := 0 to Numcrit - 1 do
    begin
    for j := 0 to length(preencherordem) - 1 do
    begin
      if peso[i] = preencherordem[j] then
      TIWUserSession(WebApplication.Data).ordemcrit[i] := auxordemcrit[j];
    end;
     end;
    end; // if prob_finalizado = true

        SetLength(TIWUserSession(WebApplication.Data).NDetalhados, Numcrit, 7); // 7 é o maior número de niveis possivel
        for i := 0 to Numcrit - 1 do
        begin
          SQL.Text := 'SELECT * FROM niveis WHERE Id_user ="' + Id_User + '" AND Id_criterio = "' + Id_criterio[i] + '"';
          ExecSQL;
          open;
          First;
          j := 0;
            while not eof do
            begin
              Ndetalhados[i,j] := FieldByName('detalhamento').Asstring;
              inc(j);
              Next;
            end;
            Close;
        end;
        SQL.Text := 'SELECT count(*) FROM operacoes WHERE Id_user ="' + Id_User + '"';
        ExecSQL;
        open;
        NumOp := FieldByName('count(*)').AsInteger;
        SetLength(TIWUserSession(WebApplication.Data).ID_Operacao, NumOp);
        SetLength(TIWUserSession(WebApplication.Data).NomeOp, NumOp);
        SetLength(TIWUserSession(WebApplication.Data).RankingFinal, NumOp);
        SetLength(TIWUserSession(WebApplication.Data).DataOP, NumOp);
        SetLength(TIWUserSession(WebApplication.Data).DataFato, NumOp);
        SetLength(TIWUserSession(WebApplication.Data).DetalhesOP, NumOp);
        close;
        SQL.Text := 'SELECT * FROM operacoes WHERE Id_user ="' + Id_User + '"';
        ExecSQL;
        open;
        First;
        i := 0;
        while not eof do
        begin
         Id_Operacao[i] := IntToStr(FieldByName('Id').AsInteger);
         NomeOp[i] := FieldByName('nome').Asstring;
         DataOp[i] := FieldByName('data').Asstring;
         DataFato[i] := FieldByName('data_fato').Asstring;
         DetalhesOP[i] := FieldByName('detalhamento_operacao').Asstring;
         if prob_finalizado = true then
         RankingFinal[i] := StrToInt(FieldByName('ranking').Asstring);
         inc(i);
         Next;
        end;
        Close;
        SetLength(MatrizConseq, NumOP, Numcrit);
        SetLength(MatrizConseqNorm, NumOP, Numcrit);
        for i := 0 to NumCrit - 1 do
          for j := 0 to NumOp - 1 do
          begin
            SQL.Text := 'SELECT * FROM consequencias WHERE Id_user ="' + Id_User + '" AND Id_criterio = "' + Id_criterio[i] + '" AND Id_operacao = "'+ TIWUserSession(WebApplication.Data).Id_Operacao[j] +'"';
            ExecSQL;
            Open;
            if FieldByName('consequencia').Asstring <> '' then
            MatrizConseq[j,i] := StrToFloat(FieldByName('consequencia').Asstring);
            if FieldByName('consequencia_norm').Asstring <> '' then
            MatrizConseqNorm[j,i] := StrToFloat(FieldByName('consequencia_norm').Asstring);
            Close;
          end;
      end;
  end;

end;

procedure TtfrmResultados.btnredefiniravdiretaClick(Sender: TObject);
var
i, j : integer;
num : double;
begin
  With TIWUserSession(WebApplication.Data) do
  begin
    Finalize(pesosavdireta);
    setlength(pesosavdireta, numcrit);
    for i := 0 to Numcrit - 1 do
    begin
      edtcritavdireta[i].Enabled := true;
      edtcritavdireta[i].text := '';
      grdinputmanual.Cell[2,i+1].Text := '';
    end;
    btnverresultadosavdireta.Enabled := False;
    btnsalvar.Enabled := True;
  end;
end;

procedure TtfrmResultados.IWButton2Click(Sender: TObject);
begin

With TiwUserSession(WebApplication.Data) do
  begin

    if cmbbxCriterios.ItemIndex = -1 then
    begin
     WebApplication.ShowMessage('Selecione um critério para continuar!');
     exit;
    end;

 {   if nomecrit[cmbbxCriterios.ItemIndex] = 'Repercussão e Extensão do Crime' then
    begin
      WebApplication.ShowMessage('Não é possível editar esse critério!');
     exit;
    end;   }

    rgnvisualizarcriterios.Visible := false;
    rgncriarnovocriterio.Visible := true;

    edtnomecrit.Text := Nomecrit[cmbbxCriterios.ItemIndex];
    if prefdir[cmbbxCriterios.ItemIndex] = 1 then
    cmbbxprefdir.ItemIndex := 1
    else
    cmbbxprefdir.ItemIndex := 0;

    if critespecial[cmbbxcriterios.ItemIndex] = true then
    begin
    grdpreencherniveis.columns[0].title := '<FONT face="Verdana" size="2"><B><P align="center">Categoria</P></B></FONT><FONT size="12"></FONT>';
   // IWLabel55.caption := '*Selecione as categorias no qual cada operação se adequa em relação ao critério selecionado:';
    end
    else
    begin
    grdpreencherniveis.columns[0].title := '<FONT face="Verdana" size="2"><B><P align="center">Níveis</P></B></FONT><FONT size="12"></FONT>';
   // IWLabel55.caption := '*Selecione a opção no qual cada operação mais se adequa em relação ao critério selecionado:';
    end;

    if niveis[cmbbxCriterios.ItemIndex] = 2 then
    cmbbxniveiscriterios.ItemIndex := 0;
    if niveis[cmbbxCriterios.ItemIndex] = 3 then
    cmbbxniveiscriterios.ItemIndex := 1;
    if niveis[cmbbxCriterios.ItemIndex] = 4 then
    cmbbxniveiscriterios.ItemIndex := 2;
    if niveis[cmbbxCriterios.ItemIndex] = 5 then
    cmbbxniveiscriterios.ItemIndex := 3;
    if niveis[cmbbxCriterios.ItemIndex] = 6 then
    cmbbxniveiscriterios.ItemIndex := 4;
    if niveis[cmbbxCriterios.ItemIndex] = 7 then
    cmbbxniveiscriterios.ItemIndex := 5;

    critsel := cmbbxCriterios.ItemIndex;
    TIWUserSession(WebApplication.Data).NovoCriterio := false;
    cmbbxniveiscriteriosChange(self);
    edtn1.Text := ndetalhados[cmbbxCriterios.ItemIndex,0];
    edtn2.Text := ndetalhados[cmbbxCriterios.ItemIndex,1];
    edtn3.Text := ndetalhados[cmbbxCriterios.ItemIndex,2];
    edtn4.Text := ndetalhados[cmbbxCriterios.ItemIndex,3];
    edtn5.Text := ndetalhados[cmbbxCriterios.ItemIndex,4];
    edtn6.Text := ndetalhados[cmbbxCriterios.ItemIndex,5];
    edtn7.Text := ndetalhados[cmbbxCriterios.ItemIndex,6];

    btncriarCriterio.Visible := False;
    btneditarcrit.Visible := true;

    end;

end;

procedure TtfrmResultados.btnverresultadosavdiretaClick(Sender: TObject);
var
i, j, k, aux : integer;
pular : array of boolean;
begin
  btnverresultadosavdireta.Enabled := False;
  btnsalvar.Enabled := True;
  DestruirElementosAvDireta;
  With TiwUserSession(WebApplication.Data) do
  begin
  for i := 0 to numcrit - 1 do
  peso[i] := pesosavdireta[i];
  Smarter;
    grdresultado.Clear;
    grdresultado.ColumnCount := 2;
    grdresultado.RowCount := NumOP+1;
    for j := 0 to grdresultado.ColumnCount - 1 do
      begin
        grdresultado.cell[0,j].Alignment := taCenter;
        grdresultado.cell[0,j].wrap := True;
        grdresultado.Cell[0,j].Font.Style := [fsbold];
        grdresultado.Cell[0,j].bgcolor := clWebGOLDENROD;
        grdresultado.Cell[0,j].Font.Color := clblack;
      end;
      grdresultado.Cell[0,0].text := 'Posição';
      grdresultado.Cell[0,1].text := 'Operação';


    k := 0;
    aux := 0;
    for i := 1 to NumOP do
    begin
      k := i - aux;
      for j := 0 to NumOP - 1 do
      begin
      if TiwUserSession(WebApplication.Data).RankingFinal[j] = i then
      begin
      grdresultado.Cell[k,0].Text := inttostr(k) + 'º';
      if grdresultado.Cell[k,1].Text = '' then
      grdresultado.Cell[k,1].Text := NomeOp[j]
      else
      begin
      grdresultado.Cell[k,1].Text :=  grdresultado.Cell[k,1].Text + ', ' + NomeOp[j];
      inc(aux);
      end;
      end;
      end;
    end;
    k := NumOP - aux;
    grdresultado.RowCount := k+1;

    txtordemcrit.Clear;
    SetLength(pular, numcrit);
    for i := 0 to Numcrit - 1 do
    pular[i] := false;


    for i := 0 to Numcrit - 1 do
    begin
    for j := 0 to Numcrit - 1 do
    begin
    if (ordemcrit[j] = i) and (pular[j] = false) then
    begin
    txtordemcrit.Lines.add(inttostr(i+1) + 'º - ' + Nomecrit[j]);
    pular[j] := true;
    //txtordemcrit.Lines.add('');
   end;
    end;
    end;
  end;
  rgninputmanual.Visible := False;
  rgnresultados.Visible := True;
  btnhasseClick(self);
end;
procedure TtfrmResultados.btnverresultadosClick(Sender: TObject);
var
i, j, k, aux : integer;
pular : array of boolean;
begin
  Smarter;
  With TiwUserSession(WebApplication.Data) do
  begin
    grdresultado.Clear;
    grdresultado.ColumnCount := 2;
    grdresultado.RowCount := NumOP+1;
    for j := 0 to grdresultado.ColumnCount - 1 do
      begin
        grdresultado.cell[0,j].Alignment := taCenter;
        grdresultado.cell[0,j].wrap := True;
        grdresultado.Cell[0,j].Font.Style := [fsbold];
        grdresultado.Cell[0,j].bgcolor := clWebGOLDENROD;
        grdresultado.Cell[0,j].Font.Color := clblack;
      end;
      grdresultado.Cell[0,0].text := 'Posição';
      grdresultado.Cell[0,1].text := 'Operação';


     k := 0;
    aux := 0;
    for i := 1 to NumOP do
    begin
      k := i - aux;
      for j := 0 to NumOP - 1 do
      begin
      if TiwUserSession(WebApplication.Data).RankingFinal[j] = i then
      begin
      grdresultado.Cell[k,0].Text := inttostr(k) + 'º';
      if grdresultado.Cell[k,1].Text = '' then
      grdresultado.Cell[k,1].Text := NomeOp[j]
      else
      begin
      grdresultado.Cell[k,1].Text :=  grdresultado.Cell[k,1].Text + ', ' + NomeOp[j];
      inc(aux);
      end;
      end;
      end;
    end;
    k := NumOP - aux;
    grdresultado.RowCount := k+1;

    txtordemcrit.Clear;
    SetLength(pular, numcrit);
    //for i := 0 to Numcrit - 1 do
    //pular[i] := false;


    for i := 0 to Numcrit - 1 do
    begin
    //for j := 0 to Numcrit - 1 do
   // begin
   // if (TiwUserSession(WebApplication.Data).ordemcrit[j] = i) and (pular[j] = false) then
    //begin
    txtordemcrit.Lines.add(inttostr(i+1) + 'º - ' + Nomecrit[ordemroc[i]]);
    //pular[j] := true;
    //txtordemcrit.Lines.add('');
  // end;
    //end;
    end;
  end;
  rgnelicitacao.Visible := False;
  rgnresultados.Visible := True;
  btnhasseClick(self);
end;

procedure TtfrmResultados.btnselecionarClick(Sender: TObject);
var
i,j, aux : integer;

begin
With TiwUserSession(WebApplication.Data) do
begin
for i := 0 to lstbxcriterios.Items.Count - 1 do
  begin
    if lstbxcriterios.Selected[i] = True then
    begin
    lstbxordemcrit.Items.Add(inttostr(lstbxordemcrit.Items.Count+1) + '- ' + auxnomecrit[i]);
    {for j := 0 to Numcrit - 1 do
    begin
      if Nomecrit[j] = lstbxcriterios.Items[i] then
      aux := i;
    end; }
    for j := 0 to grddetalhamentocrit.ColumnCount - 1 do
        begin
          grdDetalhamentoCrit.Cell[i+1,j].BGColor := clBlack;
          grdDetalhamentoCrit.Cell[i+1,j].Font.Color := clWebGAINSBORO;
        end;
    grdDetalhamentoCrit.DeleteRow(i+1);

    for j := i to length(auxnomecrit) - 2 do
    auxnomecrit[j] := auxnomecrit[j+1];
    setlength(auxnomecrit, length(auxnomecrit)-1);

    lstbxcriterios.Items.Delete(i);
    lstbxcriteriosChange(self);
     //Setlength(OrdemCrit, length(ordemcrit)+1);
     //ordemcrit[length(ordemcrit)-1] := aux;
    end;
  end;
  if lstbxcriterios.Items.Count = 0 then
  begin
    SetLength(ordemroc, numcrit);
    for I := 0 to lstbxordemcrit.Items.Count - 1 do
    begin
      for j := 0 to numcrit - 1 do
      begin
        if (inttostr(i+1) + '- ' + Nomecrit[j]) = lstbxordemcrit.Items[i] then
        TiwUserSession(WebApplication.Data).ordemroc[i] := j;
      end;
    end;
    //btnvisualizacaografica.Visible := False;
    CalculaRoc;
    lstbxordemcrit.Items.Clear;
    btnverresultados.enabled := True;
    for i := 0 to Numcrit - 1 do
    begin
      lstbxordemcrit.Items.Add(inttostr(i+1) + '- ' + Nomecrit[ordemroc[i]]+ ' (Peso = ' +
      FormatFloat('0.00',TiwUserSession(WebApplication.Data).peso[
      TiwUserSession(WebApplication.Data).ordemroc[i]]) + ')');
    end;
  end;
end;
end;

procedure TtfrmResultados.btntabelaClick(Sender: TObject);
begin
 rgnresultadotabular.visible := true;
  rgnresultadohasse.visible := false;
  if rgnresultadohasse.visible = true then
  begin
    btnhasse.BackColor := clwebsilver;
    btnhasse.backcolorto := clwebgray;
    btntabela.BackColor := clWebGoldenRod;
    btntabela.backcolorto := clwebPeachPuff;
  end
  else
  begin
    btnhasse.BackColor := clWebGoldenRod;
    btnhasse.backcolorto := clwebPeachPuff;
    btntabela.BackColor := clwebsilver;
    btntabela.backcolorto := clwebgray;
    rgnbiresultado.height := 550;
  end;
end;

procedure TtfrmResultados.btnvisualizacaograficaClick(Sender: TObject);
begin
begin
 rgnBIelicitacao.visible := false;
  rgnprocroc.visible := true;
  if rgnBIelicitacao.visible = true then
  begin
    btnvisualizacaotabular.BackColor := clwebsilver;
    btnvisualizacaotabular.backcolorto := clwebgray;
    btnvisualizacaografica.BackColor := clWebGoldenRod;
    btnvisualizacaografica.backcolorto := clwebPeachPuff;
  end
  else
  begin
    btnvisualizacaotabular.BackColor := clWebGoldenRod;
    btnvisualizacaotabular.backcolorto := clwebPeachPuff;
    btnvisualizacaografica.BackColor := clwebsilver;
    btnvisualizacaografica.backcolorto := clwebgray;
  end;
end;
BtnrestartClick(self);
end;

procedure TtfrmResultados.btnvisualizacaotabularClick(Sender: TObject);
begin
 rgnBIelicitacao.visible := True;
  rgnprocroc.visible := false;
  if rgnBIelicitacao.visible = True then
  begin
    btnvisualizacaotabular.BackColor := clwebsilver;
    btnvisualizacaotabular.backcolorto := clwebgray;
    btnvisualizacaografica.BackColor := clWebGoldenRod;
    btnvisualizacaografica.backcolorto := clwebPeachPuff;
  end
  else
  begin
    btnvisualizacaotabular.BackColor := clWebGoldenRod;
    btnvisualizacaotabular.backcolorto := clwebPeachPuff;
    btnvisualizacaografica.BackColor := clwebsilver;
    btnvisualizacaografica.backcolorto := clwebgray;
  end;
  btnreiniciarClick(self);
end;



procedure TtfrmResultados.IWLink10Click(Sender: TObject);
var
i,j,k : integer;
begin
  With TiwUserSession(WebApplication.Data) do
  begin
    With UserSession.zquery2 do
      begin
      SQL.Text := 'UPDATE usuario SET prob_finalizado ="' + 'Não' + '" WHERE Id_user =' + Id_user;
      ExecSQL;
      prob_finalizado := false;
      end;

    grddetalhamentocrit.RowCount := Numcrit+1;
    grddetalhamentocrit.ColumnCount := 3;
    lstbxcriterios.Items.Clear;

    for i := 0 to Numcrit - 1 do
    lstbxcriterios.Items.Add(Nomecrit[i]);

    for j := 0 to grddetalhamentocrit.ColumnCount - 1 do
      begin
        grddetalhamentocrit.cell[0,j].Alignment := taCenter;
        grddetalhamentocrit.cell[0,j].wrap := True;
        grddetalhamentocrit.Cell[0,j].Font.Style := [fsbold];
        grddetalhamentocrit.Cell[0,j].BGCOLOR := clWebGOLDENROD;
        grddetalhamentocrit.Cell[0,j].fONT.Color := ClBlack;
      end;
      grddetalhamentocrit.Cell[0,0].Text := 'Critério';
      //grddetalhamentocrit.Cell[0,1].Text := 'Direção de Preferência';
      grddetalhamentocrit.Cell[0,1].Text := 'Pior Desempenho';
      grddetalhamentocrit.Cell[0,2].Text := 'Melhor Desempenho';


      lstbxordemcrit.Items.Clear;
      SetLength(MaxConseq, Numcrit);
      SetLength(minconseq, numcrit);
      for i := 1 to lstbxcriterios.Items.Count do
      begin
      MaxConseq[i-1] := -9999999999;
      MinConseq[i-1] := 99999999999;
      for j := 0 to NumOp - 1 do
      begin
      if matrizconseq[j,i-1] > maxconseq[i-1] then
      maxconseq[i-1] := matrizconseq[j,i-1];
      if matrizconseq[j,i-1] < minconseq[i-1] then
      minconseq[i-1] := matrizconseq[j,i-1];

      end;
      if maxconseq[i-1] = minconseq[i-1]  then
      maxconseq[i-1] := minconseq[i-1] + 0.000001;


      grddetalhamentocrit.Cell[i,0].Text := lstbxcriterios.Items[i-1];
      grddetalhamentocrit.Cell[i,1].Text := inttostr(round(minconseq[i-1]));
      grddetalhamentocrit.Cell[i,2].Text := inttostr(Round(maxconseq[i-1]));
      end;
      lstbxcriterioROC.Clear;
  lstbxrankROC.Clear;
  //rgnROC.Visible := false;
  With TIWUserSession(WebApplication.Data) do
  begin
  for i := 0 to numcrit - 1 do
  lstbxcriterioROC.Items.Add('C' + inttostr(i+1) + ': ' + Nomecrit[i]);
  end;
      if prob_finalizado = false then
      begin
      rgnresultados.Visible := False;
      rgnelicitacao.Visible := true;
      Finalize(ordemcrit);
      gerar_grafico_overall;
      btnvisualizacaotabularClick(self);
      end;
      btnvisualizacaotabular.Visible := True;
      btnvisualizacaografica.Visible := True;

  end;

end;

procedure TtfrmResultados.IWLink11Click(Sender: TObject);
begin
rgnopcoes.Visible := true;
rgnelicitacao.Visible := false;
btnvisualizacaotabular.Visible := true;
btnvisualizacaografica.Visible := true;
BtnrestartClick(self);
btnreiniciarClick(self);
btnvisualizacaotabular.Visible := True;
btnvisualizacaografica.Visible := True;
end;

procedure TtfrmResultados.IWLink12Click(Sender: TObject);
begin
rgnelicitacao.Visible := false;
rgninputmanual.Visible := true;
criarelementosavdireta;
//btnredefiniravdiretaclick(self);
end;

procedure TtfrmResultados.IWLink13Click(Sender: TObject);
begin
rgnvisualizarcriterios.Visible := true;
rgncriarnovocriterio.Visible := false;
edtnomecrit.Clear;
cmbbxniveiscriterios.ItemIndex := -1;
edtn1.Clear;
edtn2.Clear;
edtn3.Clear;
edtn4.Clear;
edtn5.Clear;
edtn6.Clear;
edtn7.Clear;
edtn1.visible := false;
edtn2.visible := false;
edtn3.visible := false;
edtn4.visible := false;
edtn5.visible := false;
edtn6.visible := false;
edtn7.visible := false;

end;

procedure TtfrmResultados.IWLink14Click(Sender: TObject);
begin
TtfrmEntrada.Create(WebApplication).Show;
end;

procedure TtfrmResultados.IWLink1Click(Sender: TObject);
begin
rgnoperacoes.Visible := False;
rgnopcoes.Visible := true;
end;

procedure TtfrmResultados.IWLink20Click(Sender: TObject);
begin
rgninputmanual.Visible := false;
rgnelicitacao.Visible := true;
destruirelementosavdireta;
end;

procedure TtfrmResultados.IWLink23Click(Sender: TObject);
begin
WebApplication.SendFile('C:\inetpub\SADSDS\Files\Guia Rápido do Usuário - SADPF' +
      '.pdf', true);
end;



procedure TtfrmResultados.IWLink33Click(Sender: TObject);
var
i : integer;
begin
  With TiwUserSession(WebApplication.Data) do
  begin
 for i := 0 to NumOP - 1 do
    begin
      if critespecial[critsel] = true then
      chlstconsequenciaop[i].destroy;
      if critespecial[critsel] = false then
      cmbbxconsequenciaop[i].destroy;
    end;
  end;
rgnquestoescrit.visible := false;
rgncriterios.Visible := true;

end;

procedure TtfrmResultados.IWLink3Click(Sender: TObject);
begin
rgncriterios.Visible := false;
rgnopcoes.Visible := true;



end;

procedure TtfrmResultados.IWLink4Click(Sender: TObject);
begin
rgnvisualizarcriterios.Visible := false;
rgncriarnovocriterio.Visible := true;
btncriarCriterio.Visible := true;
btneditarcrit.Visible := false;
end;

procedure TtfrmResultados.IWLink5Click(Sender: TObject);
begin
rgnquestionario.Visible := False;
rgnoperacoes.Visible := True;
deletarquestionario;
end;

procedure TtfrmResultados.IWLink7Click(Sender: TObject);
begin
rgnopcoes.Visible := true;
rgnresultados.Visible := false;
end;

procedure TtfrmResultados.IWLink8Click(Sender: TObject);
var
i, j, k : integer;
aux : string;
pular : array of boolean;
begin
 With TIWUserSession(WebApplication.Data) do
 begin
   for I := 0 to Numcrit - 1 do
     grdexportacao.Columns.Add;

     grdexportacao.TotalRows := 4;
     grdexportacao.RowCount := 4;

     grdexportacao.Cells[0,0] := 'Usuário:';
     grdexportacao.Cells[0,1] := 'Data:';
     grdexportacao.Cells[1,0] :=  nome;
     aux := datetostr(date) + '. ';
     grdexportacao.Cells[1,1] := aux;

   for I := 0 to NumOP do
     grdexportacao.Columns.Add;

     for I := 0 to Numcrit do
     grdexportacao.Columns.Add;

     grdexportacao.Cells[0,3] := 'Resultados:';

     grdexportacao.TotalRows := 9 + NumOP;
     grdexportacao.RowCount := 9 + NumOP;

     grdexportacao.Cells[0,5] := 'Priorização das Operações:';
     grdexportacao.Cells[1,6] := 'Operação:';
     grdexportacao.Cells[0,6] := 'Posição:';

     for j := 1 to grdresultado.RowCount - 1 do
     begin
      grdexportacao.Cells[0,6+j] := grdresultado.Cell[j,0].Text;
      grdexportacao.Cells[1,6+j] := grdresultado.Cell[j,1].Text;
     end;

      grdexportacao.TotalRows := 12 + grdresultado.RowCount + Numcrit;
     grdexportacao.RowCount := 12 + grdresultado.RowCount + Numcrit;
     grdexportacao.Cells[0,7+grdresultado.RowCount] := 'Ordenação dos Critérios:';
     grdexportacao.Cells[0,8+grdresultado.RowCount] := 'Posição:';
     grdexportacao.Cells[1,8+grdresultado.RowCount] := 'Critério:';
     grdexportacao.Cells[2,8+grdresultado.RowCount] := 'Peso:';
     {for i := 0 to Numcrit - 1 do
     begin
      grdexportacao.Cells[1,10+NumOP+i] := Nomecrit[i];
      grdexportacao.Cells[0,10+NumOP+i] := inttostr(Ordemcrit[i]+1);
      grdexportacao.Cells[2,10+NumOP+i] := FormatFloat('0.00', peso[i]);
     end;}
     SetLength(pular, numcrit);
    for i := 0 to Numcrit - 1 do
    pular[i] := false;

    k := 0;
    for i := 0 to Numcrit - 1 do
    begin
    for j := 0 to Numcrit - 1 do
    begin
    if (ordemcrit[j] = i) and (pular[j] = false) then
    begin
      grdexportacao.Cells[1,9+grdresultado.RowCount+k] := Nomecrit[j];
      grdexportacao.Cells[0,9+grdresultado.RowCount+k] := inttostr(i+1) + 'º';
      grdexportacao.Cells[2,9+grdresultado.RowCount+k] := FormatFloat('0.00', peso[j]);
      inc(k);
    pular[j] := true;
    //txtordemcrit.Lines.add('');
   end;
    end;
    end;













    TIWAdvWebGridExcelIO1.AdvWebGrid := grdExportacao;
    TIWAdvWebGridExcelIO1.XLSExport('Relatorio Final - ' + nome +  '.xls', WebApplication);
 end;
end;

procedure TtfrmResultados.lnkalterarOPClick(Sender: TObject);
var
i : integer;
begin
With TiwUserSession(WebApplication.Data) do
begin
TIWUserSession(WebApplication.Data).OpSel := -1;
for I := 0 to TIWUserSession(WebApplication.Data).NumOP - 1 do
begin
  if grdroperacoes.RowSelect[i] = true then
  begin
  TIWUserSession(WebApplication.Data).OpSel := i;
  break;
  end;

end;
  if OpSel = -1 then
  begin
    WebApplication.ShowMessage('Selecione uma operação para continuar!');
     exit;
  end;
end;
rgnoperacoes.Visible := false;
rgnquestionario.Visible := True;
lbloperacaosel1.Caption := TIWUserSession(WebApplication.Data).NomeOP[TIWUserSession(WebApplication.Data).OpSel];
//txtpergunta.Lines[0] := 'Selecione a opção que a operação "' + lbloperacaosel1.Caption + '" mais se adequa em relação ao critério "Dano à Pessoa Humana".';
criarquestionario;
end;

procedure TtfrmResultados.lnkexcluirOPClick(Sender: TObject);
var
i,j,k : integer;
begin
With TiwUserSession(WebApplication.Data) do
begin
TIWUserSession(WebApplication.Data).OpSel := -1;
for I := 0 to TIWUserSession(WebApplication.Data).NumOP - 1 do
begin
  if grdroperacoes.RowSelect[i] = true then
  begin
  TIWUserSession(WebApplication.Data).OpSel := i;
  break;
  end;

end;
  if OpSel = -1 then
  begin
    WebApplication.ShowMessage('Selecione uma operação para continuar!');
     exit;
  end;



  With UserSession.zquery2 do
  begin
    SQL.Text := 'DELETE FROM operacoes WHERE Id=' + Id_Operacao[OpSel];
    ExecSQL;
    SQL.Text := 'DELETE FROM consequencias WHERE Id_Operacao=' + Id_Operacao[OpSel];
    ExecSQL;
  end;

  With UserSession.zquery2 do
      begin
      SQL.Text := 'UPDATE usuario SET prob_finalizado ="' + 'Não' + '" WHERE Id_user =' + Id_user;
      ExecSQL;
      prob_finalizado := false;
      end;

  grdroperacoes.DeleteSelectedRows;
  for I := OpSel to Numop - 2 do
  begin
  for j := 0 to numcrit - 1 do
  begin
    TIWUserSession(WebApplication.Data).MatrizConseq[i,j] := MatrizConseq[i+1,j];
  end;

    TIWUserSession(WebApplication.Data).NomeOP[i] := NomeOp[i+1];
    DataOp[i] := DataOp[i+1];
    DataFato[i] := datafato[i+1];
    DetalhesOP[I] := DetalhesOp[i+1];
    Id_Operacao[i] := Id_Operacao[i+1];
  end;
  NumOp := NumOp - 1;
  SetLength(MatrizConseq, NumOp, Numcrit);
  SetLength(MatrizConseqNorm,NumOp);
  SetLength(NomeOp, NumOp);
  SetLength(DataOP, Numop);
  SetLength(datafato, numop);
  SetLength(DetalhesOP, NumOP);
  SetLength(RankingFinal,numop);
  SetLength(Id_Operacao,numop);
end;
end;

procedure TtfrmResultados.lstbxcriterioROCClick(Sender: TObject);
var
  i: integer;
begin
  for i := 0 to TIWUserSession(WebApplication.Data).Numcrit - 1 do
  begin
    if TIWUserSession(WebApplication.Data).vetorGrafico[i] <> 2 then
    begin
      if i = lstbxcriterioROC.ItemIndex then
        TIWUserSession(WebApplication.Data).vetorGrafico[i] := 1
      else
        TIWUserSession(WebApplication.Data).vetorGrafico[i] := 0;
    end;
  end;
  gerar_grafico_overall;
end;

procedure TtfrmResultados.lstbxcriteriosChange(Sender: TObject);
var
i,j,k : integer;
begin
for i := 0 to lstbxcriterios.Items.Count - 1 do
  begin
    if lstbxcriterios.Selected[i] = True then
    begin
      for k := 0 to grddetalhamentocrit.RowCount - 1 do
      begin
        if lstbxcriterios.Items[i] = grddetalhamentocrit.Cell[k,0].Text then
        begin
          for j := 0 to grddetalhamentocrit.ColumnCount - 1 do
        begin
          grdDetalhamentoCrit.Cell[i+1,j].BGColor := clWebGAINSBORO;
          grdDetalhamentoCrit.Cell[i+1,j].Font.Color := clblack;
        end;
        end;

      end;

    end
    else
    begin
     for j := 0 to grddetalhamentocrit.ColumnCount - 1 do
        begin
          grdDetalhamentoCrit.Cell[i+1,j].BGColor := clBlack;
          grdDetalhamentoCrit.Cell[i+1,j].Font.Color := clWebGAINSBORO;
        end;


    end;
  end;

end;

initialization

TtfrmResultados.SetAsMainForm;

end.
