#include "Protheus.ch"
#INCLUDE "Rwmake.CH" 
#INCLUDE "tbiconn.ch"
#Define ENTER  Chr(13)+Chr(10)
/*
________________________________________________________________________________
���������������������������������������������������������������������������������
��+-----------+------------+-------+----------------------+------+------------+��
��� Fun��o    � VALDCRE    � Autor � Edson Sales          � Data � 29/09/2022 ���
��+-----------+------------+-------+----------------------+------+------------+��
��� Descri��o � Consultar se o Produto apontado na OP possui DCRE preenchida  ���
��+-----------+---------------------------------------------------------------+��
���CAMPO      � C2_PRODUTO                                                    ���
��+-----------+---------------------------------------------------------------+��
���������������������������������������������������������������������������������
���������������������������������������������������������������������������������
*/

User Function VALDCRE()
    
    Local lRet           := .T.
    Private cCodPA       := M->C2_PRODUTO
    Private cDcreVazio   := EMPTY(ALLTRIM(POSICIONE("SB1",1,XFILIAL("SB1")+cCodPA,"B1_DCRE")))
    Private cTipoPA      := POSICIONE("SB1",1,XFILIAL("SB1")+cCodPA,"B1_TIPO")=="PA"

    If  cDcreVazio .and. cTipoPA .and. RTRIM(SM0->M0_CODFIL) = '01'
        MsgStop("O produto "+LTRIM(cCodPA)+", n�o possui DCRE informado."+ENTER+"Favor procurar o setor de Gest�o do PCP!")
        //lRetorno := MsgYesNo("Deseja gerar a Ordem de produ��o como Prevista?", "Aten��o")
        
        if !MsgYesNo("Deseja gerar a Ordem de produ��o como Prevista?", "Aten��o")
           lRet := .F.
        else
            M->C2_TPOP = 'P'
        endif 
    endif

Return lRet
