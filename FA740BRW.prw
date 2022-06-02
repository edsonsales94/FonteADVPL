#INCLUDE "Rwmake.ch"
#INCLUDE "Protheus.ch"

/*_______________________________________________________________________________
���������������������������������������������������������������������������������
��+-----------+------------+-------+----------------------+------+------------+��
��� Fun��o    � FA740BRW   � Autor � Edson P. S. Sales    � Data � 02/062022  ���
��+-----------+------------+-------+----------------------+------+------------+��
��� Descri��o �   P.E. adicionar vendedor no Menu da mBrowse contas a receber ���
��+-----------+---------------------------------------------------------------+��
���������������������������������������������������������������������������������
�������������������������������������������������������������������������������*/
User Function FA740BRW()    
    Local aArea := GetArea()
    Local aRotina := {}

    aAdd(aRotina, {"Adicionar Vendedor", 'u_adicVend', 0 , 9})
    RestArea(aArea)

Return aRotina

//�����������������������������������������������������Ŀ
//�           Tela para selecionar o vendedor           �
//�������������������������������������������������������

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
        MsgAlert('Verifique se o vendendor ja est� adicionado, caso n�o esteja certifique que o titulo est� em aberto', 'N�o foi possivel adicionar o vendedor')
    
    endif

Return

//�����������������������������������������������������Ŀ
//�  Fun��o para grava na tabela o vendedor selecionado �
//�������������������������������������������������������
Static Function gravVend(cVend)
    RecLock('SE1', .F.)
        SE1->E1_VEND1 := cVend
        FWAlertSuccess('Vendedor adicionado com sucesso', 'Sucesso')
    MsUnlock()
Return 
