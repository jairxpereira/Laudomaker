#include "FiveWin.ch"

static oWnd		// Janela principal
static oDlgEdit
static oBar		// Barra de ferramentas
static obtnGo
static olstprod		// lista de produtos

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
static omtprima   // lbl de nome da matéria prima
static ousercli   // lbl de nome do cliente usuário

// controle de códigos emitidos
static gnxproduto  	// next produto
static gnxcliente 	// next cliente
static gnxoperador 	// next operador
static gnxmtprima 	// next mt prima

static lNovo		//flag de novo registro


Function Main()
local oIcon


	wakeupsys()	// Inicializa o sistema

	Set Date to British

     	// 80,95 linha,coluna
	DEFINE WINDOW oWnd TITLE sappname FROM 120,160 TO 570,800 PIXEL COLOR 0,11993179;
	MENU MenuMaker()
	DEFINE BUTTONBAR oBar SIZE 33, 33 _3D OF oWnd

	SET MESSAGE OF oWnd TO sappname DATE


        DEFINE BUTTON OF oBar FILE "..\bitmaps\client.bmp";
        MESSAGE "Cadastro de clientes" ACTION dlgCliente()

        DEFINE BUTTON OF oBar FILE "..\bitmaps\objects.bmp";
        MESSAGE "Cadastro de produtos" ACTION dlgproduto()


        DEFINE BUTTON OF oBar FILE "..\bitmaps\wndnew.bmp";
        MESSAGE "Emitir o laudo" ACTION dlgEmitir()


        DEFINE BUTTON OF oBar FILE "..\bitmaps\exit.bmp";
        MESSAGE "Sair do sistema" ACTION oWnd:End()


	ACTIVATE WINDOW oWnd

	closesys()

Return nil

Function MenuMaker()
local oMenu

	MENU oMenu
	MENUITEM  "&Cadastro"
		MENU
		MENUITEM "Clientes" ACTION dlgCliente()
		MENUITEM "Produtos" ACTION dlgproduto()
		MENUITEM "Papel" ACTION dlgmtprima()
		SEPARATOR
		MENUITEM "Operadores" ACTION dlgOper()
		MENUITEM "....." ACTION Dummy()
		SEPARATOR

		MENUITEM "Sair" ACTION oWnd:End()
		ENDMENU


	MENUITEM  "&Emissão"
		MENU
		MENUITEM "Emitir laudo" ACTION dlgEmitir()
		MENUITEM "....." ACTION dummy()
		MENUITEM "....." ACTION dummy()
		ENDMENU

	MENUITEM  "Relatórios" ACTION Dummy()
		MENU
		MENUITEM "Clientes" ACTION dummy()
		MENUITEM "Produtos" ACTION dummy()
		MENUITEM "Papel" ACTION dummy()
		ENDMENU

	MENUITEM  "Configuração" ACTION Dummy()
		MENU
		MENUITEM "configuração" ACTION dlgConfig()
		MENUITEM "....." ACTION dummy()
		MENUITEM "....." ACTION dummy()
		ENDMENU

	MENUITEM  "Ferramentas" ACTION Dummy()
		MENU
		MENUITEM "....." ACTION dummy()
		MENUITEM "....." ACTION dummy()
		MENUITEM "....." ACTION dummy()
		ENDMENU

	MENUITEM  "Intercâmbio" ACTION Dummy()
		MENU
		MENUITEM "Exportar::Cadastro" ACTION dummy()
		MENUITEM "Importar::Cadastro" ACTION dummy()
		MENUITEM "CheckUp" ACTION dummy()
		ENDMENU


	MENUITEM  "Sistema" ACTION Dummy()
		MENU
		MENUITEM "Configuração" ACTION dlgConfig()
		MENUITEM "Manual" ACTION dummy()
		MENUITEM "Auto::Verificação" ACTION dummy()
		ENDMENU

        MENUITEM  "Sai&r" ACTION oWnd:End()

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


	Define Dialog oDlgEdit FROM 0,0 TO 25,50 TITLE "Cadastro.Clientes"

	@ 1,1 SAY "Código: "  OF oDlgEdit SIZE 24,8
	@ 2,1 SAY "Nome: " OF oDlgEdit SIZE 48,8
	@ 3,1 SAY "Laudo.modelo: "  OF oDlgEdit SIZE 48,8

	@ 1.2,6 GET aoDados[1] VAR  gncdgcliente SIZE 40,10 OF oDlgEdit UPDATE
	@ 2.4,6 GET aoDados[2]  VAR gsclinome SIZE 135,10 OF oDlgEdit UPDATE
	@ 3.6,6 GET aoDados[3]  VAR gnlaudotype SIZE 40,10 OF oDlgEdit UPDATE

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

	aoDados := array(10)
	acDados := array(10)
	oCbx := array(4)

	Define Dialog oDlgEdit FROM 0,0 TO 25,75 TITLE "Cadastro.papel"

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
	@ 7.2,20 GET aoDados[3]  VAR gsgmtfront SIZE 55,10 OF oDlgEdit UPDATE // gmt do frontal
	@ 8.4,20 GET aoDados[4]  VAR gsgmtprot SIZE 55,10 OF oDlgEdit UPDATE // gmt do protetor
	@ 9.6,20 GET aoDados[4]  VAR gsrelease SIZE 55,10 OF oDlgEdit UPDATE // release

	@ 4.8,10 GET aoDados[3]  VAR gsfxesptotal SIZE 55,10 OF oDlgEdit UPDATE // fx da espessura ttl
	@ 6.0,10 GET aoDados[4]  VAR gsfxgmtades SIZE 55,10 OF oDlgEdit UPDATE // fx da gmt do adesivo
	@ 7.2,10 GET aoDados[3]  VAR gsfxgmtfront SIZE 55,10 OF oDlgEdit UPDATE // fx da gmt do front
	@ 8.4,10 GET aoDados[4]  VAR gsfxgmtprot SIZE 55,10 OF oDlgEdit UPDATE // fx da gmt do protetor
	@ 9.6,10 GET aoDados[4]  VAR gsfxrelease SIZE 55,10 OF oDlgEdit UPDATE // fx do release

        @ 8,1 BUTTON "<<" ACTION dummy() SIZE 40,10
        @ 8,8 BUTTON "&SALVAR" ACTION dummy() SIZE 40,10
	@ 8,15 BUTTON ">>" ACTION dummy() SIZE 40,10


	@ 9,1 BUTTON obtnGo PROMPT "1";
	ACTION dummy() SIZE 40,10


	@ 9,8 BUTTON "INFO" ACTION dummy() SIZE 40,10
	@ 9,15 BUTTON "SAIR" ACTION oDlgEdit:end() SIZE 40,10


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

	Define Dialog oDlgEdit FROM 0,0 TO 25,60 TITLE "Cadastro.produto"

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

	@ 1.2,9 GET aoDados[1]  VAR gncdgproduto SIZE 48,10 OF oDlgEdit UPDATE // código.produto
	@ 2.4,9 GET aoDados[2]  VAR gsmodeloprdt SIZE 96,10 OF oDlgEdit UPDATE // modelo
        @ 3.6,9 GET aoDados[2]  VAR gncdgmtprima SIZE 48,10 OF oDlgEdit UPDATE // cdg.mt.prima.usada
	@ 4.8,9 GET aoDados[3]  VAR gncdgcliente SIZE 48,10 OF oDlgEdit UPDATE // cdg.cliente.usuário
	@ 6.0,9 GET aoDados[4]  VAR gscodigoprdt SIZE 96,10 OF oDlgEdit UPDATE // cdg.datasul ou cliente
	@ 7.2,9 GET aoDados[3]  VAR gsdescrproduto SIZE 96,10 OF oDlgEdit UPDATE // descrição
	@ 8.4,9 GET aoDados[4]  VAR gsdimen SIZE 96,10 OF oDlgEdit UPDATE // dimensão
	@ 9.6,9 GET aoDados[4]  VAR gscores SIZE 96,10 OF oDlgEdit UPDATE // cores

        @ 8,1 BUTTON "<<" ACTION dummy() SIZE 40,10
        @ 8,8 BUTTON "&SALVAR" ACTION saveprdtdata() SIZE 40,10
	@ 8,15 BUTTON ">>" ACTION dummy() SIZE 40,10
	@ 8,22 BUTTON "NOVO" ACTION dummy() SIZE 40,10


	@ 9,1 BUTTON obtnGo PROMPT "1";
	ACTION dummy() SIZE 40,10


	@ 9,8 BUTTON "INFO" ACTION dummy() SIZE 40,10
	@ 9,15 BUTTON "SAIR" ACTION oDlgEdit:end() SIZE 40,10
	@ 9,22 BUTTON "LISTA" ACTION listaProduto() SIZE 40,10


	Activate Dialog oDlgEdit CENTERED
return nil

function dlgEmitir()
local aoDados, lJogado, lTerminado, lPredileto, oCbx
local sFileFullPath, sImage, sRegistro, oOldPic, oOldBtn
local sFullFileName, sGameName, sDrive, ohdSave, sMsg
local acDados, lSaving
local olblmodelo, olblcliente, olblmtprima

	aoDados := array(10)
	acDados := array(10)
	oCbx := array(4)

	Define Dialog oDlgEdit FROM 0,0 TO 35,60 TITLE "Emitindo laudo..."

	//Cliente //Produto //Material
        @ 0.8,1 BUTTON "Cdg cliente" ACTION listaClientes() SIZE 60,10

	@ 1.7,1 BUTTON "Cdg produto: " ACTION listaProduto() SIZE 60,10
        @ 2.5,1 BUTTON "Cdg mt prima usada:"  ACTION listaMtprima() SIZE 60,10

        @ 1.2,9 GET aoDados[1] VAR  gncdgcliente SIZE 26,10 OF oDlgEdit UPDATE // código do cliente
	@ 2.4,9 GET aoDados[2]  VAR gncdgproduto SIZE 26,10 OF oDlgEdit UPDATE // código do produto
        @ 3.6,9 GET aoDados[2]  VAR gncdgmtprima SIZE 26,10 OF oDlgEdit UPDATE // código da matéria prima

        @ 1,18 SAY gsclinome  OF oDlgEdit SIZE 120,8 UPDATE
	@ 2,18 SAY gsmodeloprdt OF oDlgEdit SIZE 120,8 UPDATE
        @ 3,18 SAY gsmtprimanome  OF oDlgEdit SIZE 120,8 UPDATE


       	// nota fiscal
	@ 5,1 SAY "nro.nota fiscal: "  OF oDlgEdit SIZE 48,8
	@ 5,22 SAY "data.nota fiscal: " OF oDlgEdit SIZE 48,8
	@ 6,9 GET aoDados[1] VAR  gsnfnumber SIZE 48,10 OF oDlgEdit UPDATE // código.produto
	@ 6,22 GET aoDados[2]  VAR gsnfdata SIZE 48,10 OF oDlgEdit UPDATE // modelo

        // lote rastreamento
        @ 6,1 SAY "nro.of:"  OF oDlgEdit SIZE 48,8
        @ 6,22 SAY "nro.op:"  OF oDlgEdit SIZE 96,8
        @ 7,9 GET aoDados[2]  VAR gsof SIZE 48,10 OF oDlgEdit UPDATE // cdg.mt.prima.usada
	@ 7,22 GET aoDados[3]  VAR gsop SIZE 48,10 OF oDlgEdit UPDATE // cdg.cliente.usuário

	// nro do pedido e cores
	@ 7,1 SAY "nro.pedido:"  OF oDlgEdit SIZE 96,8
	@ 8,9 GET aoDados[4]  VAR gspedido SIZE 48,10 OF oDlgEdit UPDATE // cdg.datasul ou cliente
	@ 7,22 SAY "Quantidade:"  OF oDlgEdit SIZE 96,8
	@ 8,22 GET aoDados[4]  VAR gnqtd SIZE 48,10 OF oDlgEdit UPDATE // cdg.datasul ou cliente


	@ 8.5,1 SAY "Dimensão largura x altura:"  OF oDlgEdit SIZE 96,8
	@ 10,9 GET aoDados[4]  VAR gsdimen SIZE 96,10 OF oDlgEdit UPDATE // dimensão

	@ 9.5,1 SAY "data.fabricação:"  OF oDlgEdit SIZE 48,8
	@ 9.5,22 SAY "data.validade:"  OF oDlgEdit SIZE 48,8
	@ 11,9 GET aoDados[4]  VAR gsdtfaber SIZE 48,10 OF oDlgEdit UPDATE // cores
	@ 11,22 GET aoDados[4]  VAR gsvalidade SIZE 48,10 OF oDlgEdit UPDATE // cores

	@ 10.5,1 SAY "cdg prdt datasul\cliente:"  OF oDlgEdit SIZE 96,8
	@ 12,9 GET aoDados[4]  VAR gscodigoprdt SIZE 96,10 OF oDlgEdit UPDATE // cores

	@ 11.5,1 SAY "Descrição do produto:"  OF oDlgEdit SIZE 96,8
	@ 13,9 GET aoDados[4]  VAR gsdescrproduto SIZE 96,10 OF oDlgEdit UPDATE // cores

	@ 12.2,1 SAY "cores:"  OF oDlgEdit SIZE 96,8
	@ 14,9 GET aoDados[4]  VAR gscores SIZE 96,10 OF oDlgEdit UPDATE // cores

        @ 12,1 BUTTON "<<" ACTION dummy() SIZE 40,10
        @ 12,8 BUTTON "&SALVAR" ACTION dummy() SIZE 40,10
	@ 12,15 BUTTON "EMITIR" ACTION sgetlaudo() SIZE 40,10

	@ 13,1 BUTTON obtnGo PROMPT "1"	ACTION dummy() SIZE 40,10
	@ 13,8 BUTTON "INFO" ACTION dummy() SIZE 40,10
	@ 13,15 BUTTON "SAIR" ACTION oDlgEdit:end() SIZE 40,10


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
sappname := "laudomaker 1.0.02 - Gerador de laudos"

gncdgcliente 	:= 0	//código do cliente
gncdgproduto 	:= 0 	//código do produto
gncdgmtprima 	:= 0	//código da mt prima
gncdgoper 	:= 0	//código do operador do programa

gsclinome := "AB BRASIL IND E COM DE ALIMENTOS LTDA" // nome do cliente
gsmtprimanome := "Fasson Ecogloss SP/S2045/60g"
gsdescrproduto := "Etiqueta impressa 71x51"

gscores := "PT"		//cores do produto
gsdimen	:= "71 mm x 51 mm"	//dimensão do produto
gscodigoprdt := "40020421(Jandrade's)"	//código do produto no cliente ou datasul
gsmodeloprdt := "E451 - 71x51 "     //modelo do produto para gancho de cadastro
gsesptotal := "135"	//espessura total
gsgmtades := "18,6"	//gramatura do adesivo
gsgmtfront := "84,2"	//gramatura do frontal
gsgmtprot := "60,4"	//gramtura do protetor
gsrelease := "6,4"	//release

//faixa de variação dos valores
gsfxesptotal := "135|186"	//espessura total
gsfxgmtades := "17,0|21,0"	//gramatura do adesivo
gsfxgmtfront := "82,0|88,0"	//gramatura do frontal
gsfxgmtprot := "56,0|68,2"	//gramtura do protetor
gsfxrelease := "6,0|14,0"	//release

gsnav	:= "C:\Documents and Settings\Administrador\Configurações locais\Dados de aplicativos\Google\Chrome\Application\chrome.exe"
	//path do navegador do operador
gsopernome := "Jair Pereira"	//nome do operador

gldescr	:= .t.		//vai descrição no laudo?
gldtfaber := .t.	//vai data de fabricação no laudo?
glvalidade := .t.	 //vai data de validade no laudo?
gsdtfaber := "07/08/2012" 	//data de fabricação
gsvalidade := "07/02/2012"	//data de expiração da validade
gnqtd := "10.000"		// quantidade do lote para laudo modelo 02

//configuração de códigos
gnlastoper := 1		//último código de operador cadastrado
gnlastmtprima := 1	//último código de matéria prima cadastrada
gnlastprodut := 1	//último código de produto cadastrado
gnlastcliente := 1	//último código de cliente cadastrado

//dados de coletados em runtime para emissão do laudo
gsnfdata := "18/03/2011" // data da nota fiscal
gsnfnumber := "16.800"   //número da nota fiscal
gsof := "25.669/01"	//número da of
gsop := "16.800"	//número da ordem
gspedido := "213.816"   //número do pedido

vbcrlf = chr(13) + chr(10)	//pular linha

gnlaudotype := 1
gstab := space(8)

// ponteiros de códigos
gnxproduto := 3  	// next produto
gnxcliente := 6 	// next cliente
gnxoperador := 3 	// next operador
gnxmtprima := 3 	// next mt prima

lNovo := .f.


// Abre os arquivos de trabalho
use "clientes.dbf" alias dbCliente NEW
use "mtprima.dbf" alias dbmtprima NEW
use "produto.dbf" alias dbprod NEW

ldbclientes := .t.
ldbmtprima := .t.
ldbprodutos := .t.

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
stemp := stemp + "<title>laudo</title>" + vbcrlf
stemp := stemp + "</head>" + vbcrlf
gspedido := alltrim(gspedido)
gscodigoprdt := alltrim(gscodigoprdt)


//preparando style de body
cssbody := " style=" + asp
cssbody := cssbody + "margin-top:20pt;margin-bottom:20pt;margin-left:10pt;margin-right:10pt;border-width:1px;border-style:solid"
cssbody := cssbody + asp
stemp := stemp + "<body" + cssbody + ">" + vbcrlf
stemp := stemp + "<pre>" + vbcrlf

//adicionando logotipo
stemp := stemp + "<font face='courier new' size=2>" + vbcrlf
stemp := stemp + "<center>" + vbcrlf
stemp := stemp + "<table border=1 width=90%>" + vbcrlf
stemp := stemp + "<tr>" + vbcrlf
stemp := stemp + "<td style='width:25%;margin-top:10px' valign=center>" + vbcrlf
stemp := stemp + "<img src=c:\mylab\ja_logo.jpg width=240px height=50px>" + vbcrlf
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
stemp := stemp + gstab + "Emitido por Jair Pereira" + vbcrlf
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


If !(tmpop == "")
tmpop := " - OP " + tmpop
endif

stab := space(8)

If !(tmpdescr == "")
stemp := stemp + gstab + "Descrição:      " + stab + tmpdescr + vbcrlf
endif

stemp := stemp + gstab + "Dimensões:      " + stab + gsdimen + vbcrlf
stemp := stemp + gstab + "Cor:            " + stab + gscores + " conforme padrão" + vbcrlf
stemp := stemp + gstab + "Número do Lote: " + stab + "OF " + gsof + tmpop + vbcrlf
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
stemp := stemp + gstab + "Release 180o            " + gsfxrelease + twotab + "g/1pol" + Space(3) + gsrelease + vbcrlf

stechdata := stemp + vbcrlf

return stechdata

// retorna o laudo completo
function sgetlaudo()
local slaudo

slaudo := ""
slaudo := sgetheader()
slaudo := slaudo + sgetmtprimainfo()
slaudo := slaudo + sgetprdtinfo()
slaudo := slaudo + sgetfooter()

memowrit ("c:\mylab\fivelaudo.htm", slaudo)
WinExec (gsnav + " c:\mylab\fivelaudo.htm")

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

dbCliente->(dbGoto(nrec))
gsclinome := dbCliente->SNOME
gncdgcliente := dbCliente->CDGCLIENTE

//configuração lógica
gldtfaber := dbCliente->LDATAFABER
glvalidade := dbCliente->LVALIDADE
gldescr := dbCliente->LDESCR

oDlgEdit:update()
return nil

function listaMtprima()
local nOldRec, nRecs, acMsg, ncx, sRegistro, nTotal := 0, nItem, oLbx, sData
local oDlgThis
local snome

snome := ""

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

	sRegistro := "Total de papel cadastrado #" + str(nTotal)

	//MsgList(acMsg,"Lista de Jogos Jogados")
	Define Dialog oDlgThis TITLE "Lista de papel cadastrado" FROM 0,0 TO 20,37 Style(0)

	@ 0.2,0 LISTBOX oLbx VAR nItem ITEMS acMsg SIZE 140,80
	@ 5,0 SAY sRegistro SIZE 140,10 BORDER
	@ 6.2,0 BUTTON "OK";
	ACTION (nItem := oLbx:GetPos(),setmtprimacfg(sData[nItem]),oDlgThis:End() ) SIZE 40,10
        @ 6.2,7 BUTTON "SAIR" ACTION oDlgThis:End() SIZE 40,10

        ACTIVATE DIALOG oDlgThis CENTERED
return nil

function listaProduto()
local nOldRec, nRecs, acMsg, ncx, sRegistro, nTotal := 0, nItem, oLbx, sData
local oDlgThis
local snome

	dbprod->(dbGoTop())

	nRecs := dbprod->(LastRec())

	acMsg := {}
	sData := {}


	for ncx := 1 to nRecs

	snome := dbprod->SMODELO

	sRegistro := snome

	aAdd(acMsg,sRegistro)
	aAdd(sData, dbprod->(RecNo()))
	nTotal++


	dbprod->(dbSkip(1))
	next ncx

	sRegistro := "produtos cadastrados #" + str(nTotal)


	//MsgList(acMsg,"Lista de Jogos Jogados")
	Define Dialog oDlgThis TITLE "Lista de produtos cadastrados" FROM 0,0 TO 20,37 Style(0)

	@ 0.2,0 LISTBOX oLbx VAR nItem ITEMS acMsg SIZE 140,80
	@ 5,0 SAY sRegistro SIZE 140,10 BORDER
	@ 6.2,0 BUTTON "OK";
	ACTION (nItem := oLbx:GetPos(),setprdtcfg(sData[nItem]),oDlgThis:End() ) SIZE 40,10
        @ 6.2,7 BUTTON "SAIR" ACTION oDlgThis:End() SIZE 40,10

        ACTIVATE DIALOG oDlgThis CENTERED

        oDlgEdit:update()

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
gsfxgmtprot := dbmtprima->GMTPROTFX	//gramtura do protetor
gsfxrelease := dbmtprima->RELEASEFX	//release

oDlgEdit:update()
return nil


function setprdtcfg(nrec)

dbprod->(dbGoto(nrec))

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


oDlgEdit:update()

return nil


// salva dados do cliente no arquivo
function saveclidata()
dbCliente->SNOME 	:= gsclinome
dbCliente->CDGCLIENTE 	:= gncdgcliente
dbCliente->LDATAFABER 	:= gldtfaber
dbCliente->LVALIDADE 	:= glvalidade
dbCliente->LDESCR 	:= gldescr
dbCliente->NLAUDOTYPE   := gnlaudotype
dbCliente->(dbCommit())

oDlgEdit:cTitle := "clientes - registro salvo"

return nil

// recupera arquivos do cliente no arquivo
function getclidata()
gsclinome  	:= dbCliente->SNOME
gncdgcliente  	:= dbCliente->CDGCLIENTE
gldtfaber  	:= dbCliente->LDATAFABER
glvalidade  	:= dbCliente->LVALIDADE
gldescr  	:= dbCliente->LDESCR
gnlaudotype 	:= dbCliente->NLAUDOTYPE
return nil


function saveprdtdata()

 dbprod->CDGPRDT := gncdgproduto
 dbprod->SMODELO := gsmodeloprdt
 dbprod->SCODIGO := gscodigoprdt
 dbprod->SDESCR := gsdescrproduto
 dbprod->SDIMEN := gsdimen
 dbprod->SCORES := gscores

 dbprod->(dbCommit())

 oDlgEdit:cTitle := "Produtos - registro salvo"

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

	Activate Dialog oDlgEdit CENTERED
return nil

function loadconfig

// load config
use "config.dbf" alias dbConfig NEW
dbConfig->(dbGoTop())
gnxproduto := dbConfig->NXPRODUTO
gnxcliente := dbConfig->NXCLIENTE
gnxoperador := dbConfig->NXOPERADOR
gnxmtprima := dbConfig->NXMTPRIMA
gsopernome := dbConfig->EMISSOR
gsnav := dbConfig->NAVPATH
CLOSE dbConfig
return nil

// save config
function saveconfig
use "config.dbf" alias dbConfig NEW
dbConfig->(dbGoTop())
dbConfig->NXPRODUTO := gnxproduto
dbConfig->NXCLIENTE := gnxcliente
dbConfig->NXOPERADOR := gnxoperador
dbConfig->NXMTPRIMA := gnxmtprima
dbConfig->EMISSOR := gsopernome
dbConfig->NAVPATH := gsnav
dbConfig->(dbCommit())

CLOSE dbConfig
msgInfo("Configuração salva!", "saveconfig()")
return nil


// configura caminho do navegador
function setnavpath(user)
if user == "admin"
gsnav := "C:\Documents and Settings\Administrador\Configurações locais\Dados de aplicativos\Google\Chrome\Application\chrome.exe"
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

if sdatabase == "clientes"
	if dbCliente->(bof())
	msgInfo ("começo de arquivo!",procname())
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
return nil

function btnnext (sdatabase)
local sRegistro
sRegistro := ""

if sdatabase == "clientes"
	if dbCliente->(eof())
	msgInfo ("final de arquivo!",procname())
	return nil
	endif

dbCliente->(dbSkip(1))
getclidata()
//msginfo (sdatabase + " >>")
sRegistro := str( dbCliente->(recno()) )
sRegistro := alltrim(sRegistro)
obtnGo:setText (sRegistro)
oDlgEdit:update()
endif
return nil

function btnNovo (sdatabase)
local sRegistro
sRegistro := ""

lNovo := .t.

// adiciona novo registro em clientes
if sdatabase == "clientes"

	if dbCliente->(eof())

	endif


sRegistro := str( dbCliente->(recno()) )
sRegistro := alltrim(sRegistro)

obtnGo:setText (sRegistro)
oDlgEdit:update()

endif


return nil










