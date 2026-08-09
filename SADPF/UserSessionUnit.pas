unit UserSessionUnit;

{
  This is a DataModule where you can add components or declare fields that are specific to 
  ONE user. Instead of creating global variables, it is better to use this datamodule. You can then
  access the it using UserSession.
}
interface

uses
  {IWUserSessionBase, SysUtils, Classes, DB, DBClient, WebAdapt, WebComp,
  ZConnection, ZAbstractRODataset, ZAbstractDataset, ZDataset,
  ZAbstractConnection, IWCompEdit, IWCompLabel, comobj, dateutils, IWGrids, series,IWRegion,
  IWCompMemo, IWCompListbox, IWCompText, IWTMSCheckList;}
  IWUserSessionBase, DB, DBClient, WebAdapt, WebComp,
  ZConnection, ZAbstractRODataset, ZAbstractDataset, ZDataset,
  ZAbstractConnection, Classes, SysUtils, IWAppForm, IWApplication, IWColor, IWTypes, IWCompEdit,
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
  IWCompMemo, IWAdvToolButton, ActnList;

type
  TIWUserSession = class(TIWUserSessionBase)
    ZQuery2: TZQuery;
    ZConnection2: TZConnection;
  private
    { Private declarations }
  public
    { Public declarations }
    ID_user: string;
    ID_Operacao: array of string;
    ID_criterio: array of string;
    email,senha, nome : string;
    Nomecrit, NomeOp : array of string;
    NumCrit, NumOP : integer;
    Niveis : array of integer;
    NDetalhados : array of array of string;
    PrefDir : array of integer;
    MatrizConseq : array of array of double;
    MatrizConseqNorm : array of array of double;
    RankingFinal : array of integer;
    Peso : array of double;
    DataOP, DetalhesOP : array of string;
    rgnperguntas : array of TiwRegion;
    txtperguntas : array of TiwText;
    rdgniveis : array of TiwRadioGroup;


    OpSel, critsel : integer;
    OrdemCrit, ordemROC : array of integer;
    MaxConseq, MinConseq : array of double;
    resultadoOp : array of double;
    maior, menor, valor_razao, BConseqOrd, WConseqOrd : array of real;
    vetorgrafico : array of integer;
    prob_finalizado : boolean;
    CritEspecial : array of boolean;
    lblcritAvDireta : array of tiwlabel;
    edtcritAvDireta : array of tiwedit;
    pesosavdireta : array of double;
    matrizparapar : array of array of integer;
    NumNiveis : integer;
    AuxNomeOP, levels_numalt : string;
    auxNomeCrit : array of string;
    DataFato : array of string;
    auxvalidacao : string;
    Novocriterio : boolean;
    cmbbxconsequenciaop : array of tiwcombobox;
    chlstconsequenciaop : array of ttiwchecklist;



  end;

implementation

uses untEntrada, untResultados, ServerController;

{$R *.dfm}

end.