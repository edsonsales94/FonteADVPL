#INCLUDE "totvs.ch"
#Include "RwMake.ch"
#Include "Protheus.ch"

/*
________________________________________________________________________________
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¦¦+-----------+------------+-------+----------------------+------+------------+¦¦
¦¦¦ Função    ¦ M410PVNF   ¦ Autor ¦ Edson Sales          ¦ Data ¦ 29/09/2022 ¦¦¦
¦¦+-----------+------------+-------+----------------------+------+------------+¦¦
¦¦¦ Descriçäo ¦ Ponto de entrada executado ao clicar na opção Prep. Doc Saída,¦¦¦
¦¦+-----------+---------------------------------------------------------------+¦¦
¦¦¦           ¦ para verificar a informação do DCR-e do produto acabado.      ¦¦¦
¦¦+-----------+---------------------------------------------------------------+¦¦
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
*/

User Function M410PVNF()
    Local ldShow := .T.
    Local aArea := GetArea()
    Local aAreaC5 := SC5->(GetArea())
    Local aAreaC6 := SC6->(GetArea())
    Local lRet := .T.

    PRIVATE _xdMsg :=""

    SC6->(dbSetOrder(1))
    
    ldShow := .F.
    ldCfop := .T.
    _xdMsg := " O(s) produto(s) : ( "

    //Se tiver em branco o campo, não permite prosseguir
    if SC6->(dbSeek(xFiliaL('SC5')+SC5->C5_NUM))
        While !SC6->(EoF()) .and. xFiliaL('SC5')+SC5->C5_NUM == SC6->(C6_FILIAL+C6_NUM)
           If EMPTY(ALLTRIM(POSICIONE("SB1",1,XFILIAL("SB1")+SC6->C6_PRODUTO,"B1_DCRE"))) .and. POSICIONE("SB1",1,XFILIAL("SB1")+SC6->C6_PRODUTO,"B1_TIPO")=="PA"
                _xdMsg += iIf(ldShow,",","") + SC6->C6_PRODUTO
                ldShow := .T.
            EndIf	
            if RTRIM(SC6->C6_CF) $ GetMV("MV_XCFAMOS")
               ldCfop := .F.
            endif

        SC6->(dbSkip())
        EndDo
        _xdMsg += " ), não possui DCRE informado." +Chr(13) + Chr(10)
        _xdMsg += "Favor procurar o setor de Gestão do PCP!"
        
    endif
    
 

    If ldShow .and. RTRIM(SM0->M0_CODFIL) = '01' .and. ldCfop	   
        MsgStop(_xdMsg)
        lRet := .F.
    EndIf
     
    RestArea(aAreaC6)
    RestArea(aAreaC5)
    RestArea(aArea)
Return lRet

