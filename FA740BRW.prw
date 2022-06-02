#INCLUDE "Rwmake.ch"
#INCLUDE "Protheus.ch"

/*_______________________________________________________________________________
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¦¦+-----------+------------+-------+----------------------+------+------------+¦¦
¦¦¦ Função    ¦ FA740BRW   ¦ Autor ¦ Edson P. S. Sales    ¦ Data ¦ 02/062022  ¦¦¦
¦¦+-----------+------------+-------+----------------------+------+------------+¦¦
¦¦¦ Descriçäo ¦   P.E. adicionar vendedor no Menu da mBrowse contas a receber ¦¦¦
¦¦+-----------+---------------------------------------------------------------+¦¦
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯*/
User Function FA740BRW()    
    Local aArea := GetArea()
    Local aRotina := {}

    aAdd(aRotina, {"Adicionar Vendedor", 'u_adicVend', 0 , 9})
    RestArea(aArea)

Return aRotina

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³           Tela para selecionar o vendedor           ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ

USER FUNCTION adicVend()

    Local oVend
    Local cVend := SE1->E1_VEND1

    if Empty(cVend) .and. SE1->E1_SALDO > 0

        DEFINE DIALOG oDlg TITLE "Adicionar Vendedor"  FROM 240,1 TO 440,440 PIXEL
        
        @ 10,10 TO 60,180
        oSay1  := TSay():New(20,018,{||'Selecionar vendedor'},oDlg,,,,,,.T.,CLR_BLACK,CLR_WHITE,200,12)
        @ 030,018 MSGET	oVend Var cVend	 PICTURE "@!" F3 "SA3" SIZE 100,07 PIXEL OF oDlg
        oButArq		:= TButton():New(65,110, "Confirmar" ,oDlg,{||  gravVend(cVend) },40,10,,,.F.,.T.,.F.,,.F.,,,.F.)
        oButFechar	:= TButton():New(65,160, "Fechar"    ,oDlg,{||  oDlg:End() },40,10,,,.F.,.T.,.F.,,.F.,,,.F.)

	    Activate Dialog  oDlg Centered
    else
        MsgAlert('Verifique se o vendendor ja está adicionado, caso não esteja certifique que o titulo está em aberto', 'Não foi possivel adicionar o vendedor')
    
    endif

Return

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³  Função para grava na tabela o vendedor selecionado ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
Static Function gravVend(cVend)
    RecLock('SE1', .F.)
        SE1->E1_VEND1 := cVend
        FWAlertSuccess('Vendedor adicionado com sucesso', 'Sucesso')
    MsUnlock()
Return 
