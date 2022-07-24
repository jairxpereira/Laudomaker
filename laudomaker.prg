#include "FiveWin.ch"

static oWnd		// Janela principal
static oDlgEdit
static oBar		// Barra de ferramentas
static obtnGo
static olstprod		// lista de produtos
static otxtinfo		// txt info
static txtinfo 		// info for txtinfo
static oMem		// clipboard
static oRasc		// texto de rascunho
static sRascunho

static sappname 	// nome e versão do programa

static gncdgcliente 	//código do cliente
static gsclinome	//nome do cliente
static gsmtprimanome 	//nome da matéria prima
static gncdgproduto 	//código do produto
static gncdgmtprima	//código da mt prima
static gncdgoper	//código do operador do programa
static gsdescrproduto	//descrição do produto conforme necessidade do cliente
static gnlaudotype 	// modelo do laudo 1.Corn 2.Polo

static gscores		//cores do produto
static gsdimen		//dimensão do produto
static gscodigoprdt	//código do produto no cliente ou datasul
static gsmodeloprdt     //modelo do produto para gancho de cadastro
static gsesptotal	//espessura total
static gsgmtades	//gramatura do adesivo
static gsgmtfront	//gramatura do frontal
static gsgmtprot	//gramtura do protetor
static gsrelease	//release

//faixa de variação dos valores
static gsfxesptotal	//espessura total
static gsfxgmtades	//gramatura do adesivo
static gsfxgmtfront	//gramatura do frontal
static gsfxgmtprot	//gramtura do protetor
static gsfxrelease	//release

static gsnav		//path do navegador do operador
static gsopernome	//nome do operador

static gldescr		//vai descrição no laudo?
static gldtfaber	//vai data de fabricação no laudo?
static glvalidade	//vai data de validade no laudo?
static gsdtfaber	//data de fabricação
static gsvalidade	//data de expiração da validade
static gnqtd		//quantidade do lote para laudo modelo 02

//configuração de códigos
static gnlastoper	//último código de operador cadastrado
static gnlastmtprima	//último código de matéria prima cadastrada
static gnlastprodut 	//último código de produto cadastrado
static gnlastcliente	//último código de cliente cadastrado

//dados de coletados em runtime para emissão do laudo
static gsnfdata // data da nota fiscal
static gsnfnumber //número da nota fiscal
static gsof	//número da of
static gsop	//número da ordem
static gspedido //número do pedido

static vbcrlf // chr(13) + chr(10)
static gstab // space(8)
static asp // aspas duplas chr(34)

//bandeiras de verificação de banco de dados aberto
static ldbmtprima // check if dbmtprima is open
static ldbprodutos // check if dbprodutos is open
static ldbclientes // check if dbclientes is open
static ldbemitidos // check if emitidos.dbf is open
static omtprima   // lbl de nome da matéria prima
static ousercli   // lbl de nome do cliente usuário

// controle de códigos emitidos
static gnxproduto  	// next produto
static gnxcliente 	// next cliente
static gnxoperador 	// next operador
static gnxmtprima 	// next mt prima

static lNovo		//flag de novo registro
static lGravar		//flag para gravar laudo

static lIsolaCliente	//flag para isolar cliente na listagem

static gn_Memoria	// Decora o último item de produto acessado
static gn_memPaper	// Decora a última matéria prima selecionada

// Processo de pesquisa
static gsSearching	// string sendo procurada
static oSearch		// Textbox da string sendo procurada
static oldPos		// Posição do item localizado
static cAppPath	//Caminho do programa

Function Main()
local oIcon, obmp


	wakeupsys()	// Inicializa o sistema

	sappname := " Laudomaker - Ver.: 12.01 - Julho.2015_2022 - By Jair Pereira"
	
	Set Date to ANSI

     	DEFINE WINDOW oWnd TITLE sappname FROM 50,160 TO 300,620 PIXEL COLOR 0,0x00FF00;
	MENU MenuMaker()
	DEFINE BUTTONBAR oBar SIZE 33, 33 _3D OF oWnd

	 DEFINE CLIPBOARD oMem FORMAT TEXT

	//@ 0,0 BITMAP FILENAME "..\bitmaps\infohelp.bmp"

	SET MESSAGE OF oWnd TO sappname

        DEFINE BUTTON OF oBar FILE "bitmaps\client.bmp";
        MESSAGE "Cadastro de clientes" ACTION dlgCliente()

        DEFINE BUTTON OF oBar FILE "bitmaps\objects.bmp";
        MESSAGE "Cadastro de produtos" ACTION dlgproduto()

        DEFINE BUTTON OF oBar FILE "bitmaps\code.bmp";
        MESSAGE "Cadastro de mt prima: papel" ACTION dlgmtprima()

        DEFINE BUTTON OF oBar FILE "bitmaps\wndnew.bmp";
        MESSAGE "Emitir o laudo" ACTION dlgEmitir()

        DEFINE BUTTON OF oBar FILE "bitmaps\exit.bmp";
        MESSAGE "Sair do sistema" ACTION oWnd:End()

	ACTIVATE WINDOW oWnd

	closesys()

Return nil

Function MenuMaker()
local oMenu

	MENU oMenu
	MENUITEM  "&CADASTRO"
		MENU
		MENUITEM "Clientes" ACTION dlgCliente()
		MENUITEM "Produtos" ACTION dlgproduto()
		MENUITEM "Papel" ACTION dlgmtprima()
		SEPARATOR
		MENUITEM "Operadores" ACTION dlgOper()
		SEPARATOR
		MENUITEM "Sair" ACTION oWnd:End()
		ENDMENU


	MENUITEM  "&EMISSÃO"
		MENU
		MENUITEM "Emitir laudo" ACTION dlgEmitir()
		ENDMENU


	MENUITEM  "SISTEMA" ACTION Dummy()
		MENU
		MENUITEM "Configuração" ACTION dlgConfig()
		ENDMENU

        MENUITEM  "SAI&R" ACTION oWnd:End()

	ENDMENU

Return oMenu

Function Dummy()
MsgInfo("Em desenvolvimento","Estamos contruindo...")
Return nil

function closesys
dbCloseAll()
return nil

function dlgCliente()
local aoDados, lJogado, lTerminado, lPredileto, oCbx
local sFileFullPath, sImage, sRegistro, oOldPic, oOldBtn
local sFullFileName, sGameName, sDrive, ohdSave, sMsg
local acDados, lSaving

	aoDados := array(10)
	acDados := array(10)
	oCbx := array(4)

	dbCliente->(dbgoTop())
	getclidata()


	Define Dialog oDlgEdit FROM 0,0 TO 25,50 TITLE "Cadastro de clientes"

	@ 1,1 SAY "Código: "  OF oDlgEdit SIZE 24,8
	@ 2,1 SAY "Nome: " OF oDlgEdit SIZE 48,8
	@ 3,1 SAY "Laudo.modelo: "  OF oDlgEdit SIZE 48,8

	@ 1.2,6 GET aoDados[1] VAR  gncdgcliente SIZE 40,10 OF oDlgEdit UPDATE
	@ 2.4,6 GET aoDados[2]  VAR gsclinome SIZE 135,10 OF oDlgEdit UPDATE
	@ 3.6,6 GET aoDados[3]  VAR gnlaudotype SIZE 40,10 OF oDlgEdit UPDATE

	@ 10,1 GET otxtinfo VAR txtinfo SIZE 170,10 OF oDlgEdit UPDATE READONLY COLOR 0,0xFFFF

	@ 6.2,1 CHECKBOX  oCbx[1] VAR gldtfaber PROMPT "Data de fabricacao?" OF oDlgEdit SIZE 96,8 UPDATE
	@ 7.2,1 CHECKBOX  oCbx[2] VAR glvalidade PROMPT "Data de validade?" OF oDlgEdit SIZE 96,8 UPDATE
	@ 8.2,1 CHECKBOX  oCbx[3] VAR gldescr PROMPT "Descricao?" OF oDlgEdit SIZE 96,8 UPDATE

        // @ 10.2,1 CHECKBOX  oCbx[4] VAR lSaving PROMPT "Automatic Saving" OF oDlgEdit SIZE 96,8 UPDATE

        // tipos de laudo
        @ 2.6,15.5 BUTTON "1.Modelo Corn" ACTION (gnlaudotype:=1, oDlgEdit:update()) SIZE 40,10
        @ 2.6,23 BUTTON "2.Modelo Polo" ACTION (gnlaudotype:=2, oDlgEdit:update()) SIZE 40,10

        @ 8,1 BUTTON "<<" ACTION btnprevious("clientes") SIZE 40,10
        @ 8,8 BUTTON "&SALVAR" ACTION saveclidata() SIZE 40,10
	@ 8,15 BUTTON ">>" ACTION btnnext("clientes") SIZE 40,10
	@ 8,22 BUTTON "NOVO" ACTION btnNovo("clientes") SIZE 40,10

	@ 9,1 BUTTON obtnGo PROMPT "1" ACTION dummy() SIZE 40,10 UPDATE


	@ 9,8 BUTTON "INFO" ACTION dummy() SIZE 40,10
	@ 9,15 BUTTON "SAIR" ACTION oDlgEdit:end() SIZE 40,10

        @ 9,22 BUTTON "LISTA" ACTION listaClientes() SIZE 40,10




	Activate Dialog oDlgEdit CENTERED

return nil

function dlgmtprima()
local aoDados, lJogado, lTerminado, lPredileto, oCbx
local sFileFullPath, sImage, sRegistro, oOldPic, oOldBtn
local sFullFileName, sGameName, sDrive, ohdSave, sMsg
local acDados, lSaving

	aoDados := array(12)
	acDados := array(12)
	oCbx := array(4)

	dbmtprima->(dbgotop())
	loadmtprimadata()

	Define Dialog oDlgEdit FROM 0,0 TO 27,60 TITLE "Cadastro do papel"

	@ 1,1 SAY "Código: "  OF oDlgEdit SIZE 48,8
	@ 2,1 SAY "Nome: " OF oDlgEdit SIZE 48,8

	@ 3,15 SAY "Especifição: "  OF oDlgEdit SIZE 48,8
	@ 3,30 SAY "Resultado: "  OF oDlgEdit SIZE 48,8


	@ 4,1 SAY "Espessura total:"  OF oDlgEdit SIZE 96,8
	@ 5,1 SAY "Gramatura de adesivo:"  OF oDlgEdit SIZE 96,8
	@ 6,1 SAY "Gramatura do frontal:"  OF oDlgEdit SIZE 96,8
	@ 7,1 SAY "Gramatura do protetor:"  OF oDlgEdit SIZE 96,8
	@ 8,1 SAY "Release:"  OF oDlgEdit SIZE 48,8


	@ 1.2,6 GET aoDados[1] VAR  gncdgmtprima SIZE 96,10 OF oDlgEdit UPDATE  // código
	@ 2.4,6 GET aoDados[2]  VAR gsmtprimanome SIZE 170,10 OF oDlgEdit UPDATE // nome da matéria

	@ 4.8,20 GET aoDados[3]  VAR gsesptotal SIZE 55,10 OF oDlgEdit UPDATE // espessura total
	@ 6.0,20 GET aoDados[4]  VAR gsgmtades SIZE 55,10 OF oDlgEdit UPDATE // gmt do adesivo
	@ 7.2,20 GET aoDados[5]  VAR gsgmtfront SIZE 55,10 OF oDlgEdit UPDATE // gmt do frontal
	@ 8.4,20 GET aoDados[6]  VAR gsgmtprot SIZE 55,10 OF oDlgEdit UPDATE // gmt do protetor
	@ 9.6,20 GET aoDados[7]  VAR gsrelease SIZE 55,10 OF oDlgEdit UPDATE // release

	@ 4.8,10 GET aoDados[8]  VAR gsfxesptotal SIZE 55,10 OF oDlgEdit UPDATE // fx da espessura ttl
	@ 6.0,10 GET aoDados[9]  VAR gsfxgmtades SIZE 55,10 OF oDlgEdit UPDATE // fx da gmt do adesivo
	@ 7.2,10 GET aoDados[10]  VAR gsfxgmtfront SIZE 55,10 OF oDlgEdit UPDATE // fx da gmt do front
	@ 8.4,10 GET aoDados[11]  VAR gsfxgmtprot SIZE 55,10 OF oDlgEdit UPDATE // fx da gmt do protetor
	@ 9.6,10 GET aoDados[12]  VAR gsfxrelease SIZE 55,10 OF oDlgEdit UPDATE // fx do release


	@ 11,1 GET otxtinfo VAR txtinfo SIZE 170,10 OF oDlgEdit UPDATE READONLY COLOR 0,0xFFFF

        @ 9,1 BUTTON "<<" ACTION btnPrevious("dbmtprima") SIZE 40,10
        @ 9,8 BUTTON "&SALVAR" ACTION savemtprimadata() SIZE 40,10
	@ 9,15 BUTTON ">>" ACTION btnNext("dbmtprima") SIZE 40,10


	@ 10,1 BUTTON obtnGo PROMPT "1";
	ACTION dummy() SIZE 40,10


	@ 10,8 BUTTON "INFO" ACTION dummy() SIZE 40,10
	@ 10,15 BUTTON "SAIR" ACTION oDlgEdit:end() SIZE 40,10

	@ 9,22 BUTTON "NOVO" ACTION btnNovo("dbmtprima") SIZE 40,10
  	@ 10,22 BUTTON "LISTA" ACTION listamtprima() SIZE 40,10


	Activate Dialog oDlgEdit CENTERED
return nil

function dlgproduto()
local aoDados, lJogado, lTerminado, lPredileto, oCbx
local sFileFullPath, sImage, sRegistro, oOldPic, oOldBtn
local sFullFileName, sGameName, sDrive, ohdSave, sMsg
local acDados, lSaving

	aoDados := array(10)
	acDados := array(10)
	oCbx := array(4)

	dbprod->(dbGotop())
	loadprdtdata()


	Define Dialog oDlgEdit FROM 0,0 TO 27,60 TITLE "Cadastro.produto"

	@ 1,1 SAY "Código: "  OF oDlgEdit SIZE 48,8
	@ 2,1 SAY "Modelo: " OF oDlgEdit SIZE 48,8

        @ 2.6,1 BUTTON "Cdg mt prima usada:"  ACTION (listamtprima(),oDlgEdit:update()) SIZE 60,10
	@ 3.4,1 BUTTON "cdg.cliente usuário:"  ACTION (listaClientes(),oDlgEdit:update()) SIZE 60,10

	@ 3.6,15 GET omtprima VAR gsmtprimanome OF oDlgEdit SIZE 96,8  UPDATE READONLY COLOR 0,0XFFFF
	@ 4.8,15 GET ousercli VAR gsclinome OF oDlgEdit SIZE 96,8 UPDATE READONLY COLOR 0,0XFFFF

	@ 5,1 SAY "Cdg datasul ou cliente:"  OF oDlgEdit SIZE 96,8
	@ 6,1 SAY "Descrição:"  OF oDlgEdit SIZE 96,8
	@ 7,1 SAY "Dimensão largura x altura:"  OF oDlgEdit SIZE 96,8
	@ 8,1 SAY "Cores:"  OF oDlgEdit SIZE 48,8

	@ 11.5,1 GET otxtinfo VAR txtinfo SIZE 170,10 OF oDlgEdit UPDATE READONLY COLOR 0,0xFFFF

	@ 1.2,9 GET aoDados[1]  VAR gncdgproduto SIZE 48,10 OF oDlgEdit UPDATE // código.produto
	@ 2.4,9 GET aoDados[2]  VAR gsmodeloprdt SIZE 145,10 OF oDlgEdit UPDATE // modelo
        @ 3.6,9 GET aoDados[3]  VAR gncdgmtprima SIZE 48,10 OF oDlgEdit UPDATE // cdg.mt.prima.usada
	@ 4.8,9 GET aoDados[4]  VAR gncdgcliente SIZE 48,10 OF oDlgEdit UPDATE // cdg.cliente.usuário
	@ 6.0,9 GET aoDados[5]  VAR gscodigoprdt SIZE 96,10 OF oDlgEdit UPDATE // cdg.datasul ou cliente
	@ 7.2,9 GET aoDados[6]  VAR gsdescrproduto SIZE 96,10 OF oDlgEdit UPDATE // descrição
	@ 8.4,9 GET aoDados[7]  VAR gsdimen SIZE 96,10 OF oDlgEdit UPDATE // dimensão
	@ 9.6,9 GET aoDados[8]  VAR gscores SIZE 96,10 OF oDlgEdit UPDATE // cores

        @ 9,1 BUTTON "<<" ACTION btnPrevious("dbprod") SIZE 40,10
        @ 9,8 BUTTON "&SALVAR" ACTION saveprdtdata() SIZE 40,10
	@ 9,15 BUTTON ">>" ACTION btnNext("dbprod") SIZE 40,10
	@ 9,22 BUTTON "NOVO" ACTION btnNovo("dbprod") SIZE 40,10


	@ 10,1 BUTTON obtnGo PROMPT "1";
	ACTION dummy() SIZE 40,10


	@ 10,8 BUTTON "INFO" ACTION dummy() SIZE 40,10
	@ 10,15 BUTTON "SAIR" ACTION oDlgEdit:end() SIZE 40,10
	@ 10,22 BUTTON "LISTA" ACTION listaProduto() SIZE 40,10


	Activate Dialog oDlgEdit CENTERED;
	ON RIGHT CLICK dummy()
return nil

function dlgEmitir()
local aoDados, lJogado, lTerminado, lPredileto, oCbx
local sFileFullPath, sImage, sRegistro, oOldPic, oOldBtn
local sFullFileName, sGameName, sDrive, ohdSave, sMsg
local acDados, lSaving
local olblmodelo, olblcliente, olblmtprima
local oFntNegrito, oGravar

	aoDados := array(10)
	acDados := array(10)
	oCbx := array(4)


	Define Dialog oDlgEdit FROM 0,0 TO 42,60 TITLE "Emitindo laudo..."

	Define Font oFntNegrito NAME "Courier New" SIZE 10,14 BOLD UNDERLINE

	//Cliente //Produto //Material
        @ 0.8,1 BUTTON "Cdg cliente" ACTION listaClientes() SIZE 60,10

	@ 1.7,1 BUTTON "Cdg produto: " ACTION listaProduto() SIZE 60,10
        @ 2.5,1 BUTTON "Cdg mt prima usada:"  ACTION listaMtprima() SIZE 60,10

        @ 1.2,9 GET aoDados[1] VAR  gncdgcliente SIZE 26,10 OF oDlgEdit UPDATE READONLY COLOR 0,0xd7FF90 //código do cliente
	@ 2.4,9 GET aoDados[2]  VAR gncdgproduto SIZE 26,10 OF oDlgEdit UPDATE READONLY COLOR 0,0xd7ff90  //código do produto
        @ 3.6,9 GET aoDados[2]  VAR gncdgmtprima SIZE 26,10 OF oDlgEdit UPDATE READONLY COLOR 0,0xd7ff90 // código da matéria prima

        @ 1.2,13 GET otxtinfo VAR gsclinome SIZE 122,10 OF oDlgEdit UPDATE READONLY COLOR 0,0xFFFF
        @ 2.4,13 GET otxtinfo VAR gsmodeloprdt SIZE 122,10 OF oDlgEdit UPDATE READONLY COLOR 0,0xFFFF
        @ 3.6,13 GET otxtinfo VAR gsmtprimanome SIZE 122,10 OF oDlgEdit UPDATE READONLY COLOR 0,0xFFFF

       	// nota fiscal
      	@ 5,1  SAY "Nº NF: "  OF oDlgEdit SIZE 46,8 COLOR RGB(0,0,255) FONT oFntNegrito
	@ 5,22 SAY "Data NF: " OF oDlgEdit SIZE 46,8 COLOR RGB(0,0,255) FONT oFntNegrito
	@ 6,9 GET aoDados[1] VAR  gsnfnumber SIZE 48,10 OF oDlgEdit UPDATE // código.produto
	@ 6,22 GET aoDados[2]  VAR gsnfdata SIZE 48,10 OF oDlgEdit UPDATE // modelo

        // lote rastreamento
        @ 6,1  SAY "Nº OF: "  OF oDlgEdit SIZE 48,8 COLOR RGB(0,0,255) FONT oFntNegrito
        @ 6,22 SAY "Nº OP: "  OF oDlgEdit SIZE 96,8 COLOR RGB(0,0,255) FONT oFntNegrito
        @ 7,9 GET aoDados[2]  VAR gsof SIZE 48,10 OF oDlgEdit UPDATE // cdg.mt.prima.usada
	@ 7,22 GET aoDados[3]  VAR gsop SIZE 48,10 OF oDlgEdit UPDATE // cdg.cliente.usuário

	// nro do pedido e cores
	@ 7,1 SAY "Nº Pedido:"  OF oDlgEdit SIZE 96,8 COLOR RGB(0,0,255) FONT oFntNegrito
	@ 8,9 GET aoDados[4]  VAR gspedido SIZE 48,10 OF oDlgEdit UPDATE // cdg.datasul ou cliente
	@ 7,22 SAY "Qtde:"  OF oDlgEdit SIZE 96,8 COLOR RGB(0,0,255) FONT oFntNegrito
	@ 8,22 GET aoDados[4]  VAR gnqtd SIZE 48,10 OF oDlgEdit UPDATE // cdg.datasul ou cliente


	@ 8.5,1 SAY "Dimensão largura x altura:"  OF oDlgEdit SIZE 96,8 // 96,8
	@ 10,9 GET aoDados[4]  VAR gsdimen SIZE 152,10 OF oDlgEdit UPDATE // dimensão

	@ 9.5,1 SAY "data.fabricação:"  OF oDlgEdit SIZE 48,8
	@ 9.5,22 SAY "data.validade:"  OF oDlgEdit SIZE 48,8
	@ 11,9 GET aoDados[4]  VAR gsdtfaber SIZE 48,10 OF oDlgEdit UPDATE // cores
	@ 11,22 GET aoDados[4]  VAR gsvalidade SIZE 48,10 OF oDlgEdit UPDATE // cores

	@ 10.5,1 SAY "cdg prdt datasul\cliente:"  OF oDlgEdit SIZE 96,8
	@ 12,9 GET aoDados[4]  VAR gscodigoprdt SIZE 152,10 OF oDlgEdit UPDATE // cores

	@ 11.5,1 SAY "Descrição do produto:"  OF oDlgEdit SIZE 140,8 // 96,8
	@ 13,9 GET aoDados[4]  VAR gsdescrproduto SIZE 152,10 OF oDlgEdit UPDATE // cores

	@ 12.2,1 SAY "cores:"  OF oDlgEdit SIZE 96,8 // 96,8
	@ 14,9 GET aoDados[4]  VAR gscores SIZE 152,10 OF oDlgEdit UPDATE // cores

        @ 11,1 BUTTON "COLAR.Esperto" ACTION (doParse(), oDlgEdit:update()) SIZE 41,10
        @ 11,8.5 BUTTON "SALVAR" ACTION quicksave() SIZE 40,10
	@ 11,16 BUTTON "CARREGAR" ACTION quickload() SIZE 40,10
	@ 11,23.5 BUTTON "EMITIR" ACTION sgetlaudo() SIZE 40,10
	@ 11,31 BUTTON ">>" ACTION dummy() SIZE 40,10


	@ 11.8,1 BUTTON obtnGo PROMPT "." ACTION nil SIZE 40,10
	@ 11.8,8.5 BUTTON "." ACTION nil SIZE 40,10
	@ 11.8,16 BUTTON "." ACTION nil SIZE 40,10
	@ 11.8,23.5 BUTTON "." ACTION nil SIZE 40,10
	@ 11.8,31 BUTTON "LISTA" ACTION dlgLoadLaudo() SIZE 40,10

	// Texto de rascunho
	Define Font oFntNegrito NAME "Arial" SIZE 8,14 UNDERLINE
	@ 15.3,15 SAY "Área de rascunho"  OF oDlgEdit SIZE 96,8 Font oFntNegrito // 96,8
	@ 18.5,1 GET oRasc VAR sRascunho SIZE 220,40 OF oDlgEdit TEXT

	@ 15.8,1 BUTTON obtnGo PROMPT "." ACTION nil SIZE 40,10
	@ 15.8,8.5 BUTTON "LIMPAR" ACTION oRasc:setText ( space(16000) ) SIZE 40,10
	@ 15.8,16 BUTTON "COLAR" ACTION Colar() SIZE 40,10
	@ 15.8,23.5 BUTTON "MINIMIZAR" ACTION oDlgEdit:Minimize() SIZE 40,10
	@ 15.8,31 BUTTON "SAIR" ACTION oDlgEdit:end() SIZE 40,10


	@23,1 Checkbox oGravar VAR lGravar PROMPT "Gravar laudo na emissao" OF oDlgEdit SIZE 140,12


	Activate Dialog oDlgEdit CENTERED

return nil

function dlgOper()
local aoDados, lJogado, lTerminado, lPredileto, oCbx
local sFileFullPath, sImage, sRegistro, oOldPic, oOldBtn
local sFullFileName, sGameName, sDrive, ohdSave, sMsg
local acDados, lSaving

	aoDados := array(10)
	acDados := array(10)
	oCbx := array(4)

	Define Dialog oDlgEdit FROM 0,0 TO 16,83 TITLE "Cadastro.operadores"

	@ 1,1 SAY "Código: "  OF oDlgEdit SIZE 48,8
	@ 2,1 SAY "Nome: " OF oDlgEdit SIZE 48,8
	@ 3,1 SAY "Navegador: "  OF oDlgEdit SIZE 48,8

	@ 1.2,6 GET aoDados[1] VAR  gncdgoper SIZE 40,10 OF oDlgEdit UPDATE
	@ 2.4,6 GET aoDados[2]  VAR gsopernome SIZE 110,10 OF oDlgEdit UPDATE
	@ 4.4,1 GET aoDados[3]  VAR gsnav SIZE 310,10 OF oDlgEdit UPDATE

	//@ 6.2,1 CHECKBOX  oCbx[1] VAR acDados[5] PROMPT "Data de fabricacao?" OF oDlgEdit SIZE 96,8 UPDATE
	//@ 7.2,1 CHECKBOX  oCbx[2] VAR acDados[6] PROMPT "Data de validade?" OF oDlgEdit SIZE 96,8 UPDATE
        //@ 8.2,1 CHECKBOX  oCbx[3] VAR acDados[7] PROMPT "PREDILETO" OF oDlgEdit SIZE 96,8 UPDATE
        //@ 10.2,1 CHECKBOX  oCbx[4] VAR lSaving PROMPT "Automatic Saving" OF oDlgEdit SIZE 96,8 UPDATE

        @ 4,1 BUTTON "<<" ACTION dummy() SIZE 40,10
        @ 4,8 BUTTON "&SALVAR" ACTION dummy() SIZE 40,10
	@ 4,15 BUTTON ">>" ACTION dummy() SIZE 40,10
	@ 4,22 BUTTON "NOVO" ACTION dummy() SIZE 40,10

	@ 5,1 BUTTON obtnGo PROMPT "1" ACTION dummy() SIZE 40,10


	@ 5,8 BUTTON "INFO" ACTION dummy() SIZE 40,10
	@ 5,15 BUTTON "SAIR" ACTION oDlgEdit:end() SIZE 40,10


	@ 5,22 BUTTON "LISTA" ACTION oDlgEdit:end() SIZE 40,10


	Activate Dialog oDlgEdit CENTERED
	return nil

function wakeupsys
txtinfo := "linha de mensagens"
lIsolaCliente := .t.
lGravar := .t. 		// Vamos gravar os laudos emitidos por default
cAppPath := curdrive() + ":\" + curdir();

// Área de rascunho
sRascunho := space(16000)
oRasc := nil

gn_Memoria  := 0 	// Último produto acessado
gn_memPaper := 0 	// última papel matéria-prima acessada

gncdgcliente 	:= 0	//código do cliente
gncdgproduto 	:= 0 	//código do produto
gncdgmtprima 	:= 0	//código da mt prima
gncdgoper 	:= 0	//código do operador do programa

gsclinome := space(40) // nome do cliente
gsmtprimanome := space(40)
gsdescrproduto := space(40)

gscores := space(40)	//cores do produto
gsdimen	:= space(40)	//dimensão do produto
gscodigoprdt := space(40)	//código do produto no cliente ou datasul
gsmodeloprdt := space(40)     //modelo do produto para gancho de cadastro
gsesptotal := space(40)	//espessura total
gsgmtades := space(40)	//gramatura do adesivo
gsgmtfront := space(40)	//gramatura do frontal
gsgmtprot := space(40)	//gramtura do protetor
gsrelease := space(40)	//release

//faixa de variação dos valores
gsfxesptotal := space(40)	//espessura total
gsfxgmtades := space(40)	//gramatura do adesivo
gsfxgmtfront := space(40)	//gramatura do frontal
gsfxgmtprot := space(40)	//gramtura do protetor
gsfxrelease := space(40)	//release

gsnav	:= "C:\Documents and Settings\Administrador\Configurações locais\Dados de aplicativos\Google\Chrome\Application\chrome.exe"
	//path do navegador do operador
gsopernome := "Jair Pereira"	//nome do operador

gldescr	:= .t.		//vai descrição no laudo?
gldtfaber := .t.	//vai data de fabricação no laudo?
glvalidade := .t.	 //vai data de validade no laudo?
gsdtfaber := sgetdate (date()) 	//data de fabricação
gsvalidade := space(40)	//data de expiração da validade
gnqtd := space(40)		// quantidade do lote para laudo modelo 02
gnqtd := space(10)

//configuração de códigos
gnlastoper := 1		//último código de operador cadastrado
gnlastmtprima := 1	//último código de matéria prima cadastrada
gnlastprodut := 1	//último código de produto cadastrado
gnlastcliente := 1	//último código de cliente cadastrado

//dados de coletados em runtime para emissão do laudo
gsnfdata := sgetdate( date() ) // data da nota fiscal


gsnfnumber := space(20)   //número da nota fiscal
gsof := space(20)	//número da of
gsop := space(20)	//número da ordem
gspedido := space(20)   //número do pedido

vbcrlf = chr(13) + chr(10)	//pular linha

gnlaudotype := 1
gstab := space(8)

// ponteiros de códigos
gnxproduto := 99 // next produto
gnxcliente := 99 	// next cliente
gnxoperador := 99 	// next operador
gnxmtprima := 99 	// next mt prima

// Pesquisa
gsSearching := space(40)
oldPos = 1


lNovo := .f.


// Abre os arquivos de trabalho
use "clientes.dbf" alias dbCliente NEW
use "mtprima.dbf" alias dbmtprima NEW
use "produto.dbf" alias dbprod NEW
use "emitidos" alias dbEmitidos NEW

ldbclientes := .t.
ldbmtprima := .t.
ldbprodutos := .t.
ldbemitidos := .t.

loadconfig()	// carrega a configuração


return nil

function sgetheader ()
local sheaderdata
local stemp

local scliente
local scdgproduto
local spedido
local cssbody
local tmpcliente

asp := Chr(34)


tmpcliente := ""
stemp := ""
scliente := ""
scdgproduto := ""
spedido := ""
cssbody := ""
sheaderdata := ""

//cabeçalho do arquivo html
stemp := stemp + "<html>" + vbcrlf
stemp := stemp + "<head>" + vbcrlf
stemp := stemp + "<title>" + makefilename() +"</title>" + vbcrlf
stemp := stemp + "</head>" + vbcrlf
gspedido := alltrim(gspedido)
gscodigoprdt := alltrim(gscodigoprdt)


//preparando style de body
cssbody := " style=" + asp
cssbody := cssbody + "margin-top:5pt;margin-bottom:5pt;margin-left:5pt;margin-right:5pt"
cssbody := cssbody + asp
stemp := stemp + "<body" + cssbody + ">" + vbcrlf
stemp := stemp + "<pre>" + vbcrlf

//adicionando logotipo
stemp := stemp + "<font face='courier new' size=2>" + vbcrlf
stemp := stemp + "<center>" + vbcrlf
stemp := stemp + "<table border=1 width=90%>" + vbcrlf
stemp := stemp + "<tr>" + vbcrlf
stemp := stemp + "<td style='width:25%;margin-top:10px' valign=center>" + vbcrlf
stemp := stemp + "<img src=ja_logo.jpg width=240px height=50px>" + vbcrlf
stemp := stemp + "</td>" + vbcrlf
stemp := stemp + "<td><center><b>LAUDO DE ANÁLISE PRODUTO FINAL<b></center></td>" + vbcrlf
stemp := stemp + "</tr>" + vbcrlf
stemp := stemp + "</table>" + vbcrlf
stemp := stemp + "</center>" + vbcrlf

//adicionando introdução do laudo
stemp := stemp + "<br><br>" + vbcrlf
stemp := stemp + "<center><b>CERTIFICADO DE QUALIDADE</b></center><br>" + vbcrlf + vbcrlf

tmpcliente := upper(gsclinome)
tmpcliente := alltrim (tmpcliente)

// ok
//If tmp_cliente == "" Then tmp_cliente = InputBox("Digite o nome do cliente:")

scliente := gstab + "Cliente: <b>" + tmpcliente + "</b>" + vbcrlf

spedido := gspedido
spedido := alltrim(spedido)

if !(spedido == "")
spedido := ", pedido nº " + spedido
endif

stemp := stemp + scliente + vbcrlf
stemp := stemp + "<font style='font-size:12px'>" + vbcrlf
stemp := stemp + gstab + "Certificamos que a etiqueta impressa, código " + gscodigoprdt
stemp := stemp + spedido + vbcrlf + gstab + "foi confeccionada conforme a especificação abaixo.</font>" + vbcrlf

sheaderdata := stemp + vbcrlf

return sheaderdata

function sgetfooter()
local stemp
stemp := ""

stemp := stemp + "<br><br><br><br><br><br><br><br><br><br><br><br><font style='font-size:12px'>" + vbcrlf
stemp := stemp + gstab + "Emitido por " + alltrim(gsopernome) + vbcrlf
stemp := stemp + gstab + "Depto de Qualidade - J.Andrade's Ind e Com Gráfico Ltda" + vbcrlf
stemp := stemp + gstab + "qualidade@jandrades.com.br</font>" + vbCrLf
stemp := stemp + "</pre></body></html>"
return stemp

//Configura as informações do produto desenvolvido
function sgetprdtinfo()
local  stemp
local  sprodutoinfo
local  stab
local  tmpdescr
local  tmpop

stemp := ""
tmpop := ""
sprodutoinfo := ""
stab := ""
tmpop := ""
tmpdescr := ""


tmpdescr := alltrim(gsdescrproduto)
tmpop := alltrim(gsop)
gsdtfaber := alltrim(gsdtfaber)
gsvalidade := alltrim(gsvalidade)
gsdimen := alltrim(gsdimen)
gscores := alltrim(gscores)
gsof := alltrim(gsof)
gsnfnumber := alltrim(gsnfnumber)


If !(tmpop == "")
tmpop := "OP: " + tmpop
endif

stab := space(8)

If !(tmpdescr == "")
stemp := stemp + gstab + "Descrição:      " + stab + tmpdescr + vbcrlf
endif

// 22/04/2015 - Não existe mais OF na empresa
stemp := stemp + gstab + "Dimensões:      " + stab + gsdimen + vbcrlf
stemp := stemp + gstab + "Cor:            " + stab + gscores + " conforme padrão" + vbcrlf
stemp := stemp + gstab + "Número do Lote: " + stab + tmpop + vbcrlf
stemp := stemp + gstab + "Nota Fiscal:    " + stab + gsnfnumber + " - " + gsnfdata + vbcrlf
stemp := stemp + gstab + "Medição:        " + stab + gsdimen + vbcrlf

If !(gsdtfaber == "")
stemp := stemp + gstab + "Dt fabricação:  " + stab + gsdtfaber+ vbcrlf
endif

If !(gsvalidade == "")
stemp := stemp + gstab + "Validade:       " + stab + gsvalidade + vbcrlf
End If

stemp := stemp + vbcrlf
sprodutoinfo := stemp
return sprodutoinfo


// Coleta informação da matéria prima
function sgetmtprimainfo()
local stemp
local stechdata
local twotab


stemp := ""
stechdata := ""
twotab := space(16)


stemp := stemp + gstab + "<b>Material: </b>" + gsmtprimanome + vbcrlf

stemp := stemp + gstab + "<b>Característica          Especificação   Unidade Resultado</b>" + vbcrlf

twotab := space(16 - Len(gsfxesptotal))
stemp := stemp + gstab + "Espessura total         " + gsfxesptotal + twotab + "micra" + gstab + gsesptotal + vbcrlf


twotab := space(16 - Len(gsfxgmtades))
stemp := stemp + gstab + "Gramatura de adesivo    " + gsfxgmtades + twotab + "g/m2" + gstab + gsgmtades + vbcrlf

twotab := space(16 - Len(gsfxgmtfront))
stemp := stemp + gstab + "Gramatura do frontal    " + gsfxgmtfront + twotab + "g/m2" + gstab + gsgmtfront + vbcrlf

twotab := space(16 - Len(gsfxgmtprot))
stemp := stemp + gstab + "Gramatura do protetor   " + gsfxgmtprot + twotab + "g/m2" + gstab + gsgmtprot + vbcrlf


twotab := space(16 - Len(gsfxrelease))
stemp := stemp + gstab + "Release 180o            " + gsfxrelease + twotab + "g/1pol" + Space(6) + gsrelease + vbcrlf

stechdata := stemp + vbcrlf

return stechdata

// retorna o laudo completo
function sgetlaudo()
local slaudo, sfilename
sfilename := space(40)

gnqtd := alltrim (gnqtd)

slaudo := ""

if gnqtd == ""
slaudo := sgetheader()
slaudo := slaudo + sgetmtprimainfo()
slaudo := slaudo + sgetprdtinfo()
slaudo := slaudo + sgetfooter()
else
slaudo := makehead()
slaudo := slaudo + makebody()
slaudo := slaudo + makefooter()

endif

sfilename := makefilename()
sfilename := cAppPath + "\laudos_emitidos\" + sfilename
memowrit (sfilename, slaudo)
WinExec (gsnav + sfilename)

	if (lGravar)
	GravarLaudo()
	oDlgEdit:cTitle := "Emissao de Laudo - Laudo gravado"
	endif


gnqtd := alltrim(gnqtd)
if (gnqtd == "")
gnqtd := space(20)
endif


return nil

// Lista...
Function listaClientes()
local nOldRec, nRecs, acMsg, ncx, sRegistro, nTotal := 0, nItem, oLbx, sData
local oDlgThis
local sclinome

	dbCliente->(dbGoTop())

	nRecs := dbCliente->(LastRec())

	acMsg := {}
	sData := {}

	for ncx := 1 to nRecs

	sclinome := dbCliente->SNOME

	sRegistro := sclinome

	aAdd(acMsg,sRegistro)
	aAdd(sData, dbCliente->(RecNo()))
	nTotal++

	dbCliente->(dbSkip(1))
	next ncx

	sRegistro := "Total de clientes:" + str(nTotal)

	//MsgList(acMsg,"Lista de Jogos Jogados")
	Define Dialog oDlgThis TITLE "Lista de clientes" FROM 0,0 TO 20,37 Style(0)

	@ 0.2,0 LISTBOX oLbx VAR nItem ITEMS acMsg SIZE 140,80
	@ 5,0 SAY sRegistro SIZE 140,10 BORDER
	@ 6.2,0 BUTTON "OK";
	ACTION (nItem := oLbx:GetPos(),setcliconfig(sData[nItem]),oDlgThis:End() ) SIZE 40,10
        @ 6.2,7 BUTTON "SAIR" ACTION oDlgThis:End() SIZE 40,10

        ACTIVATE DIALOG oDlgThis CENTERED


Return nil

function setcliconfig (nrec)
local sRegistro, lsetdate
sRegistro := space(10)

lsetdate := .t.

dbCliente->(dbGoto(nrec))
gsclinome := dbCliente->SNOME
gncdgcliente := dbCliente->CDGCLIENTE

//configuração lógica
gldtfaber := dbCliente->LDATAFABER
glvalidade := dbCliente->LVALIDADE
gldescr := dbCliente->LDESCR

// data de validade == .f.?
if glvalidade == .f.
gsdtfaber := space(20)
gsvalidade := space(20)
endif

// data de validade == .t.?

// Corn e Ab Brasil não tem data de validade no laudo
// if (gncdgcliente == 1) .or. (gncdgcliente ==2)
// msgInfo("Não tem data de validade nem dt de fabricação no laudo desse cliente",gsclinome)
// lsetdate := .f.
// endif

// Mude validade da Omamori para 18 meses
if (gncdgcliente == 8)
gsvalidade := "18 meses"
lsetdate := .f.
endif

// Mude validade da AB Brasil 12 meses
if (gncdgcliente == 1)
gsvalidade := "12 meses a partir da data de fabricação"
lsetdate := .f.
msgInfo("Não preencha QTD neste cliente!!! Ok Cristina?",gsclinome)


endif


// Mude validade da KPSG para 6 meses
if (gncdgcliente == 6)
gsvalidade := "6 meses"
lsetdate := .f.
endif

// nexans := "seis meses"
if (gncdgcliente == 7)
gsvalidade := "seis meses"
msginfo (gsvalidade)
lsetdate := .f.
endif

if (glvalidade == .t. .and. lsetdate == .t.)
gsdtfaber := sgetdate(date())
gsvalidade := sgetdate(date() + 180)
endif

if (glvalidade == .f. .and. lsetdate == .t.)
gsdtfaber := space(20)
gsdtfaber := sgetdate(date())
gsvalidade := space(20)
endif


sRegistro := str( dbCliente->(recno()) )
sRegistro := alltrim(sRegistro)
if !(obtnGo == nil)
obtnGo:setText (sRegistro)
oDlgEdit:update()
endif

oDlgEdit:update()
return nil

function listaMtprima()
local nOldRec, nRecs, acMsg, ncx, sRegistro, nTotal := 0, nItem, oLbx, sData
local oDlgThis, oFont
local snome

snome := ""

	Define Font oFont NAME "Courier" SIZE 8, 10 BOLD

	dbmtprima->(dbGoTop())

	nRecs := dbmtprima->(LastRec())

	acMsg := {}
	sData := {}

	for ncx := 1 to nRecs

	snome := dbmtprima->SDESCR

	sRegistro := snome

	aAdd(acMsg,sRegistro)
	aAdd(sData, dbmtprima->(RecNo()))
	nTotal++


	dbmtprima->(dbSkip(1))
	next ncx

	sRegistro := "TOTAL DE PAPÉIS CADASTRADOS #" + str(nTotal)


	Define Dialog oDlgThis TITLE "LISTA DE PAPÉIS CADASTRADOS" FROM 0,0 TO 30, 55 Style(0) Font oFont COLOR 0xff0000,0xeeeeee

	@ 0.2,0 LISTBOX oLbx VAR nItem ITEMS acMsg SIZE 200,160 COLOR 0xff0000,0xFFFFFF
	@ 10.5,1 SAY sRegistro SIZE 200,12 BORDER

	@ 10,1 BUTTON "Seleção Anterior";
	ACTION mpsetposition( oLbx ) SIZE 80,20

	@ 10, 19 BUTTON "OK";
	ACTION (nItem := oLbx:GetPos(),setmtprimacfg(sData[nItem]),oDlgThis:End() ) SIZE 40,20

        @  10,27.2 BUTTON "SAIR" ACTION oDlgThis:End() SIZE 40,20

        ACTIVATE DIALOG oDlgThis CENTERED

        gn_memPaper := nItem;


return nil

function listaProduto()
local nOldRec, nRecs, acMsg, ncx, sRegistro, nTotal := 0, nItem, oLbx, sData
local oDlgThis, oFont, ntempCliente
local snome

	dbprod->(dbGoTop())

	nItem := 0

	nRecs := dbprod->(LastRec())

	acMsg := {}
	sData := {}


	for ncx := 1 to nRecs

	snome := dbprod->SMODELO
	ntempCliente := dbprod->CDGCLIENTE

	// pegue apenas o produto do cliente selecionado
	if ( gncdgcliente > 0 .and. gncdgcliente == ntempCliente)
	sRegistro := snome
	aAdd(acMsg,sRegistro)
	aAdd(sData, dbprod->(RecNo()))
	nTotal++
	endif

	// pegue todos os produtos para código de cliente zerado
	if ( gncdgcliente == 0)
	sRegistro := snome
	aAdd(acMsg,sRegistro)
	aAdd(sData, dbprod->(RecNo()))
	nTotal++
	endif

	dbprod->(dbSkip(1))
	next ncx

	sRegistro := "produtos cadastrados #" + str(nTotal)

	Define Font oFont NAME "Courier" SIZE 8,10

	//MsgList(acMsg,"Lista de Jogos Jogados")
	Define Dialog oDlgThis TITLE "Lista de produtos cadastrados" FROM 0,0 TO 38,55 Style(0) FONT oFont COLOR 0xff0000,0xeeeeee


	@ 0.2,1 LISTBOX oLbx VAR nItem ITEMS acMsg SIZE 200,160
	@ 10.5,1 SAY sRegistro SIZE 200,12 BORDER

	@ 10,1 BUTTON "Seleção Anterior";
	ACTION setposition( oLbx ) SIZE 80,20

        @ 10,19 BUTTON "SAIR" ACTION (oldPos :=1, oDlgThis:End()) SIZE 40,20

        @ 10,27.2 BUTTON "OK";
	ACTION (nItem := oLbx:GetPos(),setprdtcfg(sData[nItem]),oDlgThis:End() ) SIZE 40,20


	// Processo de pesquisa
	@ 13, 1 BUTTON "Pesquisar >> ";
	ACTION (nItem := oLbx:GetPos(), searchProductDown( oLbx )) SIZE 60,20

	@ 13, 25 BUTTON "<< Pesquisar";
	ACTION (nItem := oLbx:GetPos(), searchProductUp( oLbx )) SIZE 55,20

	@ 14.5,1 SAY "Pesquisa de subtexto:" SIZE 200,12

	@ 17.8, 8.6 GET oSearch VAR  gsSearching SIZE 80,15 OF oDlgThis UPDATE

        ACTIVATE DIALOG oDlgThis CENTERED

        gn_Memoria := nItem

        oDlgEdit:update()

return nil


// Configura posição no listbox
function setposition( oLbx )
if ( gn_Memoria == 0); MsgInfo("Esse botão vai lembrar o último item selecionado!"); endif
if ( gn_Memoria != 0); oLbx:Select ( gn_Memoria ); endif
if ( gn_Memoria != 0); oLbx:SetSel ( gn_Memoria ); endif
if ( gn_Memoria != 0); oLbx:refresh ( gn_Memoria ); endif
return nil


// Configura posição no listbox de matéria-prima ( papel )
function mpsetposition( oLbx )
if ( gn_memPaper == 0); MsgInfo("Esse botão vai lembrar o último item selecionado!"); endif
if ( gn_memPaper != 0); oLbx:Select ( gn_memPaper ); endif
if ( gn_memPaper != 0); oLbx:SetSel ( gn_memPaper ); endif
if ( gn_memPaper != 0); oLbx:refresh ( gn_memPaper ); endif
return nil

function setmtprimacfg(nrec)

dbmtprima->(dbGoto(nrec))

gsmtprimanome := dbmtprima->SDESCR

gncdgmtprima := dbmtprima->CDGMTPRIMA

gsesptotal := dbmtprima->ESPTOTAL	//espessura total
gsgmtades := dbmtprima->GMTADES	//gramatura do adesivo
gsgmtfront := dbmtprima->GMTFRONT	//gramatura do frontal
gsgmtprot := dbmtprima->GMTPROT	//gramtura do protetor
gsrelease := dbmtprima->RELEASE	//release

//faixa de variação dos valores
gsfxesptotal := dbmtprima->ESPTOTALFX	//espessura total
gsfxgmtades := dbmtprima->GMTADESFX	//gramatura do adesivo
gsfxgmtfront := dbmtprima->GMTFRONTFX	//gramatura do frontal
gsfxgmtprot := dbmtprima->GMTPROTFX	//gramatura do protetor
gsfxrelease := dbmtprima->RELEASEFX	//release

oDlgEdit:update()
return nil


function setprdtcfg(nrec)
local sRegistro
sRegistro := space(10)

dbprod->(dbGoto(nrec))
gncdgmtprima := dbprod->CDGMTPRIMA
gncdgcliente := dbprod->CDGCLIENTE
lgetmtprimanome(gncdgmtprima)
lgetclinome(gncdgcliente)

gncdgproduto := dbprod->CDGPRDT
gsmodeloprdt := dbprod->SMODELO
gscodigoprdt := dbprod->SCODIGO
gsdescrproduto := dbprod->SDESCR
gsdimen := dbprod->SDIMEN
gscores := dbprod->SCORES

// filtrar depois pelo nome do cliente
// configurar o material

// gncdgmtprima := dbprod->CDGMTPRIMA
// gncdgcliente := dbprod->CDGCLIENTE

sRegistro := str( dbprod->(recno()) )
sRegistro := alltrim(sRegistro)
if !(obtnGo == nil)
obtnGo:setText (sRegistro)
oDlgEdit:update()
endif


oDlgEdit:update()

return nil


// Essa função pesquisa por um produto
function searchProductDown( oList )
local ncx, npos, ntam, stemp

ncx := 0
npos := 0
ntam := 0

ntam := oList:Len()
stemp := space(40)

stemp = alltrim ( gsSearching )


for ncx = oldPos to ntam
// MsgInfo ( oList:GetItem(ncx), " oList:GetItem(ncx)")

npos := at( stemp, oList:GetItem(ncx) )

if (npos > 0)
oList:Select ( ncx )
oldPos := ncx + 1
return nil
endif

next ncx

return nil


// Essa função pesquisa por um produto
function searchProductUp( oList )
local ncx, npos, ntam, stemp

ncx := 0
npos := 0
ntam := 0

if (oldPos < 1)
oldPos := 1
endif


ntam := oList:Len()
stemp := space(40)

stemp = alltrim ( gsSearching )

for ncx = oldPos to 1 step -1

npos := at( stemp, oList:GetItem(ncx) )

if (npos > 0)
oList:Select ( ncx )
oldPos := ncx - 1
return nil
endif

next ncx

return nil



// salva dados do cliente no arquivo
function saveclidata()

if lNovo == .t.
dbCliente->(dbAppend())
ndxupdate()
lNovo := .f.
endif

dbCliente->SNOME 	:= gsclinome
dbCliente->CDGCLIENTE 	:= gncdgcliente
dbCliente->LDATAFABER 	:= gldtfaber
dbCliente->LVALIDADE 	:= glvalidade
dbCliente->LDESCR 	:= gldescr
dbCliente->NLAUDOTYPE   := gnlaudotype
dbCliente->(dbCommit())

txtinfo := "clientes - registro salvo"
oDlgEdit:update()

return nil

// recupera arquivos do cliente no arquivo
function getclidata()
local sRegistro, ntemp
sRegistro := space(10)

gsclinome  	:= dbCliente->SNOME
gncdgcliente  	:= dbCliente->CDGCLIENTE
gldtfaber  	:= dbCliente->LDATAFABER
glvalidade  	:= dbCliente->LVALIDADE
gldescr  	:= dbCliente->LDESCR
gnlaudotype 	:= dbCliente->NLAUDOTYPE

if (gncdgcliente == 7)
gsvalidade := "6 meses"
endif

return nil


// salva dados da tabela de produtos
function saveprdtdata()

if lNovo == .t.
dbprod->(dbAppend())
ndxupdate()
lNovo := .f.
endif

dbprod->CDGPRDT := gncdgproduto
dbprod->SMODELO := gsmodeloprdt
dbprod->SCODIGO := gscodigoprdt
dbprod->SDESCR := gsdescrproduto
dbprod->SDIMEN := gsdimen
dbprod->SCORES := gscores
dbprod->CDGCLIENTE := gncdgcliente
dbprod->CDGMTPRIMA := gncdgmtprima

// irrelevante - isso vai ser pego da tabela de clientes
//dbprod->LDATAFABER := gldtfaber
//dbprod->LVALIDADE := glvalidade
//dbprod->LDESCR := gldescr

 dbprod->(dbCommit())

 txtinfo := "Produtos - registro salvo"

 oDlgEdit:update()

return nil

function dlgConfig
local aoDados, lJogado, lTerminado, lPredileto, oCbx
local sFileFullPath, sImage, sRegistro, oOldPic, oOldBtn
local sFullFileName, sGameName, sDrive, ohdSave, sMsg
local acDados, lSaving

	aoDados := array(10)
	acDados := array(10)
	oCbx := array(4)

	Define Dialog oDlgEdit FROM 0,0 TO 18,75 TITLE "Configuração"

	@ 1,1 SAY "Emissor:"  OF oDlgEdit SIZE 48,8
	@ 2,1 SAY "Caminho do navegador: " OF oDlgEdit SIZE 96,8

	//@ 3,1 SAY "nxproduto: "  OF oDlgEdit SIZE 48,8
	//@ 3,20 SAY "nxcliente: "  OF oDlgEdit SIZE 48,8

	@ 4,1 SAY "nxmtprima:"  OF oDlgEdit SIZE 96,8
	@ 5,1 SAY "nxoperador:"  OF oDlgEdit SIZE 96,8
	@ 6,1 SAY "nxproduto:"  OF oDlgEdit SIZE 96,8
	@ 7,1 SAY "nxcliente:"  OF oDlgEdit SIZE 96,8
	//@ 8,1 SAY "reservado:"  OF oDlgEdit SIZE 48,8


	@ 1.2,6 GET aoDados[1] VAR  gsopernome SIZE 96,10 OF oDlgEdit UPDATE  // emissor
	@ 3.6,1 GET aoDados[2]  VAR gsnav SIZE 270,10 OF oDlgEdit UPDATE // nav path

	//@ 3.6,18 GET aoDados[3]  VAR gnxcliente SIZE 55,10 OF oDlgEdit READONLY COLOR 0,0xffff // nxcliente
	//@ 3.6,6 GET aoDados[3]  VAR gnxproduto SIZE 55,10 OF oDlgEdit READONLY COLOR 0,0xffff  // nx produto
	@ 4.8,6 GET aoDados[3]  VAR gnxmtprima SIZE 55,10 OF oDlgEdit READONLY COLOR 0,0xffff // nx mt prima
	@ 6.0,6 GET aoDados[4]  VAR gnxoperador SIZE 55,10 OF oDlgEdit READONLY COLOR 0,0xffff // nx
	@ 7.2,6 GET aoDados[3]  VAR gnxproduto SIZE 55,10 OF oDlgEdit READONLY COLOR 0,0xffff // fx da gmt do front
	@ 8.4,6 GET aoDados[4]  VAR gnxcliente SIZE 55,10 OF oDlgEdit READONLY COLOR 0,0xffff // fx da gmt do protetor
	//@ 9.6,6 GET aoDados[4]  VAR gsfxrelease SIZE 55,10 OF oDlgEdit UPDATE // fx do release

        @ 1.8,11 BUTTON "admin" ACTION setnavpath("admin") SIZE 40,10
        @ 3.5,19 BUTTON "&SALVAR" ACTION saveconfig() SIZE 40,10
	@ 1.8,19 BUTTON "jandrades" ACTION setnavpath("jandrades") SIZE 40,10
	@ 4.3,19 BUTTON "SAIR" ACTION oDlgEdit:end() SIZE 40,10
	@ 5.8,19 BUTTON "RESET" ACTION setnavpath("reset") SIZE 40,10

	Activate Dialog oDlgEdit CENTERED
return nil

function loadconfig

// load config
use "config.dbf" alias dbConfig NEW
use "cfglocal.dbf" alias dbCfgLocal NEW

dbConfig->(dbGoTop())
gnxproduto := dbConfig->NXPRODUTO
gnxcliente := dbConfig->NXCLIENTE
gnxoperador := dbConfig->NXOPERADOR
gnxmtprima := dbConfig->NXMTPRIMA

gsopernome := dbCfgLocal->EMISSOR
gsnav := dbCfgLocal->NAVPATH
CLOSE dbConfig
CLOSE dbCfgLocal

return nil

// save config
function saveconfig
use "config.dbf" alias dbConfig NEW
use "cfglocal.dbf" alias dbCfgLocal NEW


dbConfig->(dbGoTop())
dbConfig->NXPRODUTO := gnxproduto
dbConfig->NXCLIENTE := gnxcliente
dbConfig->NXOPERADOR := gnxoperador
dbConfig->NXMTPRIMA := gnxmtprima

dbCfgLocal->EMISSOR := gsopernome
dbCfgLocal->NAVPATH := gsnav

dbConfig->(dbCommit())

CLOSE dbConfig
CLOSE dbCfgLocal

msgInfo("Configuração salva!", "saveconfig()")
return nil


// configura caminho do navegador

function setnavpath(user)

if user == "admin"
gsnav := "C:\Documents and Settings\Administrador\Configurações locais\Dados de aplicativos\Google\Chrome\Application\chrome.exe"
oDlgEdit:update()
endif

if user == "reset"
gsnav := space(255)
oDlgEdit:update()
endif


if user == "jandrades"
gsnav := "C:\Arquivos de programas\Google\Chrome\Application\chrome.exe"
oDlgEdit:update()
endif
return nil

function btnprevious (sdatabase)
local sRegistro
sRegistro := ""
lNovo := .f.

//tabela de mt prima: papel
if sdatabase == "dbmtprima"
txtinfo := "<< dbmtprima::papel"
	if dbmtprima->(bof())
	txtinfo := "dbmtprima::papel::início de arquivo!"
	oDlgEdit:update()
	return nil
	endif

dbmtprima->(dbSkip(-1))
loadmtprimadata()
sRegistro := str( dbmtprima->(recno()) )
sRegistro := alltrim(sRegistro)
obtnGo:setText (sRegistro)
oDlgEdit:update()
endif

if sdatabase == "clientes"
txtinfo := "<< cliente"
	if dbCliente->(bof())
	txtinfo := "cliente::início de arquivo!"
	oDlgEdit:update()
	return nil
	endif

dbCliente->(dbSkip(-1))
getclidata()
//msgInfo ("<< " + sdatabase)
sRegistro := str( dbCliente->(recno()) )
sRegistro := alltrim(sRegistro)
obtnGo:setText (sRegistro)
oDlgEdit:update()
endif

// tabela de produtos
if sdatabase == "dbprod"
txtinfo := "<< produtos"
	if dbprod->(bof())
	txtinfo := "produtos::início!"
	oDlgEdit:update()
	return nil
	endif

dbprod->(dbSkip(-1))
loadprdtdata()
sRegistro := str( dbprod->(recno()) )
sRegistro := alltrim(sRegistro)
obtnGo:setText (sRegistro)
oDlgEdit:update()
endif



return nil

function btnnext (sdatabase)
local sRegistro
sRegistro := ""
lNovo := .f.

//tabela de mt prima: papel
if sdatabase == "dbmtprima"
txtinfo := ">> dbmtprima::papel"
	if dbmtprima->(eof())
	txtinfo := "dbmtprima::papel::Final de arquivo!"
	oDlgEdit:update()
	return nil
	endif

dbmtprima->(dbSkip(1))
loadmtprimadata()
sRegistro := str( dbmtprima->(recno()) )
sRegistro := alltrim(sRegistro)
obtnGo:setText (sRegistro)
oDlgEdit:update()
endif


// tabela de clientes
if sdatabase == "clientes"
txtinfo := ">> cliente"
	if dbCliente->(eof())
	txtinfo := "cliente::Final de arquivo!"
	oDlgEdit:update()
	return nil
	endif

dbCliente->(dbSkip(1))
getclidata()
sRegistro := str( dbCliente->(recno()) )
sRegistro := alltrim(sRegistro)
obtnGo:setText (sRegistro)
oDlgEdit:update()
endif

// tabela de produtos
if sdatabase == "dbprod"
txtinfo := ">> produtos"
	if dbprod->(eof())
	txtinfo := "produtos::Final de arquivo!"
	oDlgEdit:update()
	return nil
	endif

dbprod->(dbSkip(1))
loadprdtdata()
sRegistro := str( dbprod->(recno()) )
sRegistro := alltrim(sRegistro)
obtnGo:setText (sRegistro)
oDlgEdit:update()
endif




return nil

function btnNovo (sdatabase)
local sRegistro, stam
sRegistro := "*"
ndxload()
txtinfo := "Novo registro << ou >> para cancelar"
obtnGo:setText (sRegistro)
oDlgEdit:update()

lNovo := .t.

// adiciona novo registro na tab de mt prima: papel
if (sdatabase == "dbmtprima")


	gnxmtprima := gnxmtprima + 1
	gncdgmtprima := gnxmtprima

 	gsmtprimanome := space(40)
 	gsesptotal := space(12)
 	gsgmtades := space(12)
 	gsgmtfront := space(12)
 	gsgmtprot := space(12)
 	gsrelease := space(12)
 	gsfxesptotal := space(12)
 	gsfxgmtades := space(12)
 	gsfxgmtfront := space(12)
 	gsfxgmtprot := space(12)
 	gsfxrelease := space(12)

	oDlgEdit:update()

endif



// adiciona novo registro em clientes
if (sdatabase == "clientes")

	gsclinome  	:= space(40)
	gnxcliente      := gnxcliente + 1
	gncdgcliente  	:= gnxcliente
	gldtfaber  	:= .f.
	glvalidade  	:= .f.
	gldescr  	:= .f.
	gnlaudotype 	:= 1
	oDlgEdit:update()

endif


// adiciona novo registro em produtos
if (sdatabase == "dbprod")

	gnxproduto := gnxproduto + 1
	gncdgproduto    := gnxproduto
	//gncdgcliente  	:= 0
	//gsmodeloprdt 	:= space(40)
	gscodigoprdt 	:= space(40)
	//gsdescrproduto 	:= space(40)
	//gsdimen 	:= space(40)
	//gscores 	:= space(40)

	oDlgEdit:update()

endif



return nil

// atualiza índices
function ndxupdate()
use "config.dbf" alias dbConfig NEW
dbConfig->(dbGoTop())
dbConfig->NXPRODUTO := gnxproduto
dbConfig->NXCLIENTE := gnxcliente
dbConfig->NXOPERADOR := gnxoperador
dbConfig->NXMTPRIMA := gnxmtprima
dbConfig->(dbCommit())
CLOSE dbConfig
return nil

// carrega índices de controle de emissão de códigos
function ndxload()
use "config.dbf" alias dbConfig NEW
dbConfig->(dbGoTop())
gnxproduto := dbConfig->NXPRODUTO
gnxcliente := dbConfig->NXCLIENTE
gnxoperador := dbConfig->NXOPERADOR
gnxmtprima := dbConfig->NXMTPRIMA
CLOSE dbConfig
return nil

// carrega dados da tabela de produtos
function loadprdtdata

gncdgproduto := dbprod->CDGPRDT
gsmodeloprdt := dbprod->SMODELO
gscodigoprdt := dbprod->SCODIGO
gsdescrproduto := dbprod->SDESCR
gsdimen := dbprod->SDIMEN
gscores := dbprod->SCORES

gncdgcliente := dbprod->CDGCLIENTE
lgetclinome(gncdgcliente)

gncdgmtprima := dbprod->CDGMTPRIMA
lgetmtprimanome(gncdgmtprima)

return nil


// localiza nome da mt prima
function lgetmtprimanome (ncdg)
local ntam, snome, cdg, lfound, ncx

ntam := 0
snome := space(80)
cdg := 0
lfound := .f.
ncx := 0

dbmtprima->(dbGotop())
ntam := dbmtprima->(lastrec())

for ncx := 1 to ntam

cdg := dbmtprima->CDGMTPRIMA

	if (cdg == ncdg)
	gsmtprimanome := dbmtprima->SDESCR
	gncdgmtprima := dbmtprima->CDGMTPRIMA
	gsesptotal := dbmtprima->ESPTOTAL
	gsgmtades := dbmtprima->GMTADES
	gsgmtfront := dbmtprima->GMTFRONT
	gsgmtprot := dbmtprima->GMTPROT
	gsrelease := dbmtprima->RELEASE
	gsfxesptotal := dbmtprima->ESPTOTALFX
	gsfxgmtades := dbmtprima->GMTADESFX
	gsfxgmtfront := dbmtprima->GMTFRONTFX
	gsfxgmtprot := dbmtprima->GMTPROTFX
	gsfxrelease := dbmtprima->RELEASEFX


	lfound := .t.
	exit
	endif

dbmtprima->(dbSkip())

next ncx

if !(oDlgEdit == nil)
oDlgEdit:update()
endif


return lfound

// localiza nome do cliente
function lgetclinome(ncdg)
// gsclinome dbCliente

local ntam, snome, cdg, lfound, ncx

ntam := 0
snome := space(80)
cdg := 0
lfound := .f.

dbCliente->(dbGotop())
ntam := dbCliente->(lastrec())

for ncx := 1 to ntam

cdg := dbCliente->CDGCLIENTE

	if (cdg == ncdg)
	gsclinome := dbCliente->SNOME
	lfound := .t.
	exit
	endif

dbCliente->(dbSkip())

next ncx
return lfound


// carrega dados técnicos do papel
function loadmtprimadata()
gncdgmtprima := dbmtprima->CDGMTPRIMA
gsmtprimanome := dbmtprima->SDESCR
gsesptotal := dbmtprima->ESPTOTAL
gsgmtades := dbmtprima->GMTADES
gsgmtfront := dbmtprima->GMTFRONT
gsgmtprot := dbmtprima->GMTPROT
gsrelease := dbmtprima->RELEASE
gsfxesptotal := dbmtprima->ESPTOTALFX
gsfxgmtades := dbmtprima->GMTADESFX
gsfxgmtfront := dbmtprima->GMTFRONTFX
gsfxgmtprot := dbmtprima->GMTPROTFX
gsfxrelease := dbmtprima->RELEASEFX
return nil

// salva dados técnicos do papel
function savemtprimadata()
// for saving
if lNovo == .t.
dbmtprima->(dbAppend())
ndxupdate()
lNovo := .f.
endif

dbmtprima->CDGMTPRIMA := gncdgmtprima
dbmtprima->SDESCR     := gsmtprimanome
dbmtprima->ESPTOTAL   := gsesptotal
dbmtprima->GMTADES    := gsgmtades
dbmtprima->GMTFRONT   := gsgmtfront
dbmtprima->GMTPROT    := gsgmtprot
dbmtprima->RELEASE    := gsrelease
dbmtprima->ESPTOTALFX := gsfxesptotal
dbmtprima->GMTADESFX  := gsfxgmtades
dbmtprima->GMTFRONTFX := gsfxgmtfront
dbmtprima->GMTPROTFX  := gsfxgmtprot
dbmtprima->RELEASEFX  := gsfxrelease

dbmtprima->(dbCommit())
txtinfo := "Produtos - registro salvo"
oDlgEdit:update()
return nil


// transforma uma data numa string decente
function sgetdate(dtinfo)
local smonth, sday, syear, sdatainfo
local nmonth, nday, nyear
nmonth := month(dtinfo)
nday := day(dtinfo)
nyear := year(dtinfo)
smonth := alltrim(str(nmonth))
sday := alltrim(str(nday))
syear := alltrim(str(nyear))
if nday <= 9
sday := "0" + alltrim(str(nday))
endif
if nmonth <= 9
smonth := "0" + alltrim(str(nmonth))
endif
sdatainfo := sday + "/" + smonth + "/" + syear
return sdatainfo


// faz o cabeçalho do laudo tipo 02
function makehead()
local stemp, tmp_lote
tmp_lote := ""
stemp := ""
// cabeçalho básico do arquivo html e do laudo
stemp := stemp + "<html>" + vbcrlf
stemp := stemp + "<head>" + vbcrlf
stemp := stemp + "<title>" + makefilename() + "</title>" + vbcrlf
stemp := stemp + "</head>" + vbcrlf
stemp := stemp + "<body style='margin-top:5pt;margin-bottom:5pt;margin-left:5pt;margin-right:5pt;'>" + vbcrlf
stemp := stemp + "<center>" + vbcrlf
stemp := stemp + vbcrlf
stemp := stemp + "<table style='margin-top:20pt;margin-bottom:20pt;margin-left:10pt;margin-right:10pt;border-width:1px;border-style:solid'>" + vbcrlf
stemp := stemp + "<tr><td style='width:25%;margin-top:10px' valign=center>" + vbcrlf
stemp := stemp + "<img src=ja_logo.jpg width=240px height=50px></td>" + vbcrlf
stemp := stemp + "<td><center><b>LAUDO DE ANÁLISE PRODUTO FINAL<b></center></td>" + vbcrlf
stemp := stemp + "</tr></table></center>" + vbcrlf
stemp := stemp + vbcrlf

// colocando dados essenciais do cabeçalho
stemp := stemp + "<center><table width=97% style='margin-top:20pt;margin-bottom:20pt;margin-left:10pt;margin-right:10pt;border-width:1px;border-style:solid'>" + vbcrlf
stemp := stemp + "<tr><td>Cliente: " + gsclinome + "</td> <td>Data: " + gsnfdata + "</td></tr>" + vbcrlf
stemp := stemp + "<tr><td>Produto: " + gsdescrproduto + "</td><td>" + "Pedido: " + gspedido + "</td></tr>" + vbcrlf

// prepara informação de lote
gsof := alltrim(gsof)
gsop := alltrim(gsop)
if !(gsof == "")
tmp_lote := "OF " + gsof
endif
if !(gsop == "")
tmp_lote := tmp_lote + " - OP " + gsop
 if (gsof == "")
 tmp_lote := "OP: " + gsop
 endif
endif

//tira espaços em branco dos campos
gnqtd := alltrim(gnqtd)
gsdtfaber := alltrim(gsdtfaber)
gsnfnumber := alltrim(gsnfnumber)
stemp := stemp + "<tr><td>Lote: " + tmp_lote + "</td> <td>Data de fabricação: " + gsdtfaber + "</td></tr>" + vbcrlf
stemp := stemp + "<tr><td>Nota Fiscal: " + gsnfnumber + "</td>  <td>Quantidade: " + gnqtd + " un</td></tr>" + vbcrlf
stemp := stemp + "</table></center>" + vbcrlf
stemp := stemp + vbcrlf
return stemp

// faz o corpo do laudo tipo 02
function makebody()
local stemp
stemp := "<center><table border=1 width=97% "
stemp := stemp + "style='margin-top:20pt;margin-bottom:20pt;margin-left:10pt;margin-right:10pt;border-width:1px;border-style:solid'>" + vbcrlf
stemp := stemp + "<tr bgcolor=#ebebeb><td colspan=3>Amostragem conforme ABNT NBR 5426:  Simples normal</td></tr>" + vbcrlf
stemp := stemp + "<tr bgcolor=#ebebeb><td>Quantidade amostrada Nível I=</td> <td colspan=2>Nível S3 =</td></tr>" + vbcrlf
stemp := stemp + "<tr bgcolor=#ebebeb><td><b><center>Teste</center></b></td>" + vbcrlf
stemp := stemp + "<td><b><center>Especificação</center></b></td>" + vbcrlf
stemp := stemp + "<td><b><center>Resultado</center></b></td></tr>" + vbcrlf
stemp := stemp + "<tr bgcolor=#ebebeb><td>Análise Visual Nível = I NQA =</td><td></td><td></td>  </tr>" + vbcrlf
stemp := stemp + vbcrlf

// remover chr(32) dessas variáveis
gsmtprimanome := alltrim(gsmtprimanome)
gscodigoprdt := alltrim(gscodigoprdt)
stemp := stemp + "<tr><td>Material</td> <td></td> <td>" + gsmtprimanome + "</td></tr>" + vbcrlf
stemp := stemp + "<tr><td>Identificação</td> <td></td> <td>" + gscodigoprdt + "</td></tr>" + vbcrlf


stemp := stemp + "<tr><td>Embalagem/Acondicionamento</td> <td></td> <td>caixa de papelão</td></tr>" + vbcrlf
stemp := stemp + "<tr><td>Sentido do embobinamento</td> <td></td> <td>Pé</td></tr>" + vbcrlf
stemp := stemp + "<tr><td>Numeração do Liner</td> <td></td> <td>Não tem</td></tr>" + vbcrlf
stemp := stemp + "<tr><td>Texto</b></td> <td></td> <td>Conforme</td></tr>" + vbcrlf
stemp := stemp + "<tr><td>Impressão geral</td> <td></td>" + vbcrlf
stemp := stemp + "<td>Nítida, sem manchas, falhas, borrões,<br>" + vbcrlf
stemp := stemp + "mau recobrimento, livre de sujidades</td></tr>" + vbcrlf

// remover chr(32) de gscores gsdimen gsesptotal
gscores := alltrim(gscores)
gsdimen := alltrim(gsdimen)
gsesptotal := alltrim(gsesptotal)

stemp := stemp + "<tr><td>Padrão de  cor</td> <td></td> <td>" + gscores + "</td></tr>" + vbcrlf
stemp := stemp + vbcrlf
stemp := stemp + "<tr bgcolor=#ebebeb><td>Análise dimensional - Nível = I - NQA =</td><td></td><td></td>  </tr>" + vbcrlf
stemp := stemp + "<tr><td>largura x altura</td> <td>+/- 1 mm</td> <td>" + gsdimen + "</td></tr>" + vbcrlf
stemp := stemp + "<tr><td>Espessura</td> <td>micras +/- 10%</td> <td>" + gsesptotal + " g/m2 </td></tr>" + vbcrlf
stemp := stemp + vbcrlf

stemp := stemp + "<tr bgcolor=#ebebeb><td>Análise Física - Nível = S3 - NQA = </td><td></td><td></td>  </tr>" + vbcrlf


// remover chr(32)
gsgmtprot := alltrim(gsgmtprot)
gsgmtfront := alltrim(gsgmtfront)
gsgmtades := alltrim(gsgmtades)
gsrelease := alltrim(gsrelease)
stemp := stemp + "<tr><td>Gramatura do Liner</td> <td> g/m2 +/- 5%</td> <td>" + gsgmtprot +  " g/m2 </td></tr>" + vbcrlf
stemp := stemp + "<tr><td>Gramatura do Frontal</td> <td> g/m2 +/- 5%</td> <td>" + gsgmtfront + " g/m2 </td></tr>" + vbcrlf
stemp := stemp + "<tr><td>Gramatura do Adesivo</td> <td> g/m2 +/- 5%</td> <td>" + gsgmtades + " g/m2 </td></tr>" + vbcrlf
stemp := stemp + "<tr><td>Tack </td> <td>N/m</td> <td>Min 600 N/m</td></tr>" + vbcrlf
stemp := stemp + "<tr><td>Adesividade</td> <td>N/m</td> <td>Min 800 N/m</td></tr>" + vbcrlf
stemp := stemp + "<tr><td>Release</td> <td>g/1pol</td> <td>" + gsrelease + " g/1pol</td></tr>" + vbcrlf
stemp := stemp + "</table></center>" + vbcrlf
stemp := stemp + vbcrlf


return stemp



// fecha o rodapé do laudo tipo 02
function makefooter()
local stemp
// remove chr(32) das variáveis
gsopernome := alltrim (gsopernome)
gsnfdata := alltrim (gsnfdata)

stemp := "<center><table border=1 width=97% style='margin-top:20pt;margin-bottom:20pt;margin-left:10pt;margin-right:10pt;border-width:1px;border-style:solid'>"
stemp := stemp + "<tr><td>Responsável: <br>" + gsopernome + "</td>" + vbcrlf
stemp := stemp + "<td>Laudo final:<br> [x] aprovado [ ] rejeitado</td>" + vbcrlf
stemp := stemp + "<td>Data:<br>" + gsnfdata + "</td></tr>" + vbcrlf
stemp := stemp + "</table></center>" + vbcrlf
stemp := stemp + vbcrlf

stemp := stemp + "</body></html>" + vbcrlf
return stemp


function makefilename()
local sfilename, nf_info, cdg_info, ntam
local letter, tempword, badletter, nbadtam, badword
local ncx, mcx

ncx := 0
mcx := 0

gsnfnumber := alltrim (gsnfnumber)
gscodigoprdt = alltrim (gscodigoprdt)
tempword := ""
letter := ""
badletter := ""


badword := ".()/\´', "

nbadtam := len(badword)
ntam = len (gsnfnumber)

// limpe o nome da nota fiscal
for ncx = 1 to ntam
	letter := substr(gsnfnumber, ncx,1)

		for mcx = 1 to nbadtam
		badletter := substr(badword, mcx,1)
			if (letter == badletter)
			letter := ""
			endif
		next
	tempword := tempword + letter
next

nf_info := "NF" + tempword
tempword := ""


// limpe o código
ntam = len (gscodigoprdt)
for ncx = 1 to ntam
	letter := substr(gscodigoprdt, ncx,1)

		for mcx = 1 to nbadtam
		badletter := substr(badword, mcx,1)
			if (letter == badletter)
			letter := ""
			endif
		next
	tempword := tempword + letter
next


//-----------------------------------------
cdg_info := tempword
tempword := ""

sfilename := nf_info + "-" + cdg_info + ".htm"


return sfilename


Function quicksave()

use "tmplaudo.dbf" alias dbtmpLaudo NEW

dbtmpLaudo->CDGPRDT 	:= gncdgproduto
dbtmpLaudo->SMODELO 	:= gsmodeloprdt
dbtmpLaudo->SCODIGO 	:= gscodigoprdt
dbtmpLaudo->SDESCR 	:= gsmtprimanome
dbtmpLaudo->SDIMEN 	:= gsdimen
dbtmpLaudo->SCORES 	:= gscores
dbtmpLaudo->SDATAFABER 	:= gsdtfaber
dbtmpLaudo->SVALIDADE 	:= gsvalidade
dbtmpLaudo->SDESCPROD 	:= gsdescrproduto
dbtmpLaudo->CDGCLIENTE 	:= gncdgcliente
dbtmpLaudo->CDGMTPRIMA 	:= gncdgmtprima
dbtmpLaudo->SPEDIDO 	:= gspedido
dbtmpLaudo->SNF  	:= gsnfnumber
dbtmpLaudo->SOP  	:= gsop
dbtmpLaudo->SOF  	:= gsof
dbtmpLaudo->SQTDE 	:= alltrim (gnqtd)

// Salva dados de matéria prima
dbtmpLaudo->CDGMTPRIMA := gncdgmtprima
dbtmpLaudo->SDESCR     := gsmtprimanome
dbtmpLaudo->ESPTOTAL   := gsesptotal
dbtmpLaudo->GMTADES    := gsgmtades
dbtmpLaudo->GMTFRONT   := gsgmtfront
dbtmpLaudo->GMTPROT    := gsgmtprot
dbtmpLaudo->RELEASE    := gsrelease
dbtmpLaudo->ESPTOTALFX := gsfxesptotal
dbtmpLaudo->GMTADESFX  := gsfxgmtades
dbtmpLaudo->GMTFRONTFX := gsfxgmtfront
dbtmpLaudo->GMTPROTFX  := gsfxgmtprot
dbtmpLaudo->RELEASEFX  := gsfxrelease

dbtmpLaudo->SCLINOME  := gsclinome
dbtmpLaudo->SNFDATA  := gsnfdata

oDlgEdit:cTitle := "Emitindo laudo - Gravação rápida Ok"


dbtmpLaudo->(dbCloseArea())

return nil



Function quickload()

use "tmplaudo.dbf" alias dbtmpLaudo NEW

gncdgproduto := dbtmpLaudo->CDGPRDT
gsmodeloprdt := dbtmpLaudo->SMODELO
gscodigoprdt := dbtmpLaudo->SCODIGO
gsmtprimanome  := dbtmpLaudo->SDESCR
gsdimen := dbtmpLaudo->SDIMEN
gscores := dbtmpLaudo->SCORES
gsdtfaber := dbtmpLaudo->SDATAFABER
gsvalidade := dbtmpLaudo->SVALIDADE
gsdescrproduto := dbtmpLaudo->SDESCPROD
gncdgcliente := dbtmpLaudo->CDGCLIENTE
gncdgmtprima := dbtmpLaudo->CDGMTPRIMA
gspedido := dbtmpLaudo->SPEDIDO
gsnfnumber := dbtmpLaudo->SNF
gsop := dbtmpLaudo->SOP
gsof := dbtmpLaudo->SOF
gnqtd := dbtmpLaudo->SQTDE

// Salva dados de matéria prima
gncdgmtprima := dbtmpLaudo->CDGMTPRIMA
gsmtprimanome := dbtmpLaudo->SDESCR
gsesptotal := dbtmpLaudo->ESPTOTAL
gsgmtades := dbtmpLaudo->GMTADES
gsgmtfront := dbtmpLaudo->GMTFRONT
gsgmtprot := dbtmpLaudo->GMTPROT
gsrelease := dbtmpLaudo->RELEASE
gsfxesptotal := dbtmpLaudo->ESPTOTALFX
gsfxgmtades := dbtmpLaudo->GMTADESFX
gsfxgmtfront := dbtmpLaudo->GMTFRONTFX
gsfxgmtprot := dbtmpLaudo->GMTPROTFX
gsfxrelease := dbtmpLaudo->RELEASEFX

gsclinome := dbtmpLaudo->SCLINOME
gsnfdata := dbtmpLaudo->SNFDATA

oDlgEdit:update()

dbtmpLaudo->(dbCloseArea())

oDlgEdit:cTitle := "Emitindo laudo - Carregamento rápido Ok"

return nil


Function GravarLaudo()

dbEmitidos->(dbAppend())

dbEmitidos->CDGPRDT 	:= gncdgproduto
dbEmitidos->SDATAFABER 	:= gsdtfaber
dbEmitidos->SVALIDADE 	:= gsvalidade
dbEmitidos->SNFDATA  := gsnfdata
dbEmitidos->CDGCLIENTE 	:= gncdgcliente
dbEmitidos->CDGMTPRIMA 	:= gncdgmtprima

dbEmitidos->SPEDIDO 	:= gspedido

gsnfnumber := alltrim ( gsnfnumber)
gscodigoprdt := alltrim( gscodigoprdt)
gsdescrproduto := alltrim( gsdescrproduto)


dbEmitidos->SNF  	:= gsnfnumber
dbEmitidos->SOP  	:= gsop
dbEmitidos->SOF  	:= gsof
dbEmitidos->SQTDE 	:= alltrim (gnqtd)

dbEmitidos->SNOMELAUDO 	:= "NF" + gsnfnumber + "-" + gsnfdata + "-" + gscodigoprdt + "-" + gsdescrproduto


return nil

Function dlgLoadLaudo()
local nOldRec, nRecs, acMsg, ncx, sRegistro, nTotal := 0, nItem, oLbx, sData
local oDlgThis, oFont, ntempCliente
local snome

	dbEmitidos->(dbGoTop())

	nRecs := dbEmitidos->(LastRec())

	acMsg := {}
	sData := {}


	for ncx := 1 to nRecs

	snome := dbEmitidos->SNOMELAUDO
	ntempCliente := dbEmitidos->CDGCLIENTE

	// pegue apenas o produto do cliente selecionado
	if ( gncdgcliente > 0 .and. gncdgcliente == ntempCliente)
	sRegistro := snome
	aAdd(acMsg,sRegistro)
	aAdd(sData, dbEmitidos->(RecNo()))
	nTotal++
	endif

	// pegue todos os produtos para código de cliente zerado
	if ( gncdgcliente == 0)
	sRegistro := snome
	aAdd(acMsg,sRegistro)
	aAdd(sData, dbEmitidos->(RecNo()))
	nTotal++
	endif

	dbEmitidos->(dbSkip(1))
	next ncx

	sRegistro := "Laudos emitidos deste cliente #" + alltrim (str(nTotal))

	Define Font oFont NAME "Courier" SIZE 8,10


	Define Dialog oDlgThis TITLE "Laudos emitidos" FROM 0,0 TO 30,75 Style(0) FONT oFont COLOR 0xff0000,0xc082ff


	@ 0.2,1 LISTBOX oLbx VAR nItem ITEMS acMsg SIZE 280,160
	@ 10.5,1 SAY sRegistro SIZE 200,12 BORDER

	@ 10,1 BUTTON "OK";
	ACTION (nItem := oLbx:GetPos(),xsetprdtcfg(sData[nItem]),oDlgThis:End() ) SIZE 40,20

        @ 10,8 BUTTON "SAIR" ACTION oDlgThis:End() SIZE 40,20

        ACTIVATE DIALOG oDlgThis CENTERED

        oDlgEdit:update()



return nil




function xsetprdtcfg(nrec)
local sRegistro, ntam, ncx
sRegistro := space(10)

ntam := 0
ncx := 0

dbEmitidos->(dbGoto(nrec))
gncdgmtprima := dbEmitidos->CDGMTPRIMA
gncdgcliente := dbEmitidos->CDGCLIENTE
gncdgproduto := dbEmitidos->CDGPRDT
lgetmtprimanome(gncdgmtprima)
lgetclinome(gncdgcliente)

gsof := dbEmitidos->SOF
gsop := dbEmitidos->SOP
gnqtd := dbEmitidos->SQTDE
gsvalidade := dbEmitidos->SVALIDADE
gsdtfaber := dbEmitidos->SDATAFABER

gsnfdata := dbEmitidos->SNFDATA    // data da nota fiscal
gsnfnumber := dbEmitidos->SNF  //número da nota fiscal
gspedido := dbEmitidos->SPEDIDO  //número do pedido


gnqtd := alltrim(gnqtd)
if (gnqtd == "")
gnqtd := space(20)
endif


// Mostra nro do registro do laudo no botão
sRegistro := str( dbEmitidos->(recno()) )
sRegistro := alltrim(sRegistro)
if !(obtnGo == nil)
obtnGo:setText (sRegistro)
oDlgEdit:update()
endif



// Localiza dados do produto e matéria prima
ntam := dbProd->(recCount())


dbProd->(dbGotop())
for ncx = 1 to ntam
	if (gncdgproduto == dbProd->CDGPRDT)
	loadprdtdata()
	exit
	endif
	dbProd->(dbSkip(1))
next ncx

oDlgEdit:update()

return nil

// Cola clipboard na área de rascunho
Function colar()
local cText
ACTIVATE CLIPBOARD oMem
cText := oMem:GetText()
oRasc:SetText (cText)
sRascunho := cText
return nil

// Fazendo leitura e colagem inteligente de informações da memória


// Realiza o parse da OP/OG
Function doParse()
local sRastreio
ACTIVATE CLIPBOARD oMem
sRastreio := oMem:GetText()

if (alltrim(sRastreio) == "")
return nil
endif

gsop := getOP(sRastreio)
gsof := getOF(sRastreio)

return nil

Function getOP(sRastreio)
local stext, ntam, ncx, sletter, stwo, ntext
local np1, np2, ofound, opfound, offound, sword
local ntamtext

sword := ""
ofound = 0
opfound := .f.
offound := .f.
sletter := ""
np1 := 0
np2 := 0

ntam = len(sRastreio)

for ncx = 1 to ntam
sletter := substr( sRastreio, ncx, 1)

	if (sletter == "O")
	stwo := substr(sRastreio, ncx+1,1)
		if (stwo == "P")
		np1 := ncx + 1
		opfound := .t.
		endif
	endif

	if (opfound == .t.)
		if (sletter == "F")
		np2 := ncx - 2
		opfound := .f.
		endif
	endif

	if (opfound == .t.)
		if (ncx == ntam)
		np2 := ntam
		endif
	endif
next ncx

ntamtext := np2 - np1
stext := substr( sRastreio, np1+1, ntamtext)
stext := alltrim (stext)

// Purifica o texto encontrado
ntam := len(stext)
for ncx = 1 to ntam
sletter := substr(stext, ncx, 1)
if (sletter == "." .or. sletter == "-" .or. sletter == " " .or. sletter == ":")
sletter := ""
endif
sword := sword + sletter
next ncx
stext := sword


return stext


Function getOF(sRastreio)
local stext, ntam, ncx, sletter, stwo, ntext
local np1, np2, ofound, opfound, offound, sword
local ntamtext

sword := ""
ofound = 0
opfound := .f.
offound := .f.
sletter := ""
np1 := 0
np2 := 0

ntam = len(sRastreio)

for ncx = 1 to ntam
sletter := substr( sRastreio, ncx, 1)

	if (sletter == "O")
	stwo := substr(sRastreio, ncx+1,1)
		if (stwo == "F")
		np1 := ncx + 1
		opfound := .t.
		endif
	endif

	if (opfound == .t.)
		if (sletter == "P")
		np2 := ncx - 2
		opfound := .f.
		endif
	endif

	if (opfound == .t.)
		if (ncx == ntam)
		np2 := ntam
		endif
	endif
next ncx

ntamtext := np2 - np1
stext := substr( sRastreio, np1+1, ntamtext)
stext := alltrim (stext)

// Purifica o texto encontrado
ntam := len(stext)
for ncx = 1 to ntam
sletter := substr(stext, ncx, 1)
if (sletter == "." .or. sletter == "-" .or. sletter == " " .or. sletter == ":")
sletter := ""
endif
sword := sword + sletter
next ncx
stext := sword

return stext


function lastproduct()
msgInfo("Vamos ir para o último produto selecionado!")
return nil























