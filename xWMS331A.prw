#INCLUDE "Protheus.ch"
#INCLUDE "Totvs.ch"
#INCLUDE "FwMvcDef.ch"
#include "RwMake.ch"

/*________________________________________________________________________________________________
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¦¦+-----------+---------------------------+-------+----------------------+------+--------------+¦¦
¦¦¦ Função    ¦ WMSA331A()      		  ¦ Autor ¦ Edson Sales          ¦ Data ¦ 17/11/2021   ¦¦¦
¦¦+-----------+---------------------------+-------+----------------------+------+--------------+¦¦
¦¦¦ Descriçäo ¦ Ponto de entrada MVC roda na rotina Monitor de Serviço WMS  			       ¦¦¦
¦¦+-----------+--------------------------------------------------------------------------------+¦¦
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯*/
User Function WMSA331A()

Local aParam      := PARAMIXB
Local xRet        := .T.
Local oObj        := ''
Local cIdPonto    := ''
Local cIdModel    := ''
Local lIsGrid     := .F.
Local nPeso       := 0
Local nOpcA,nI, nX := 1
Local cProduto
Local cDoc   	
Local cSerie 	
Local cCliFor	
Local cLoja	
Local cServ	
Local cTar	
Local cAtiv
Local nVal     
 
If aParam <> NIL
    oObj       := aParam[1]
    cIdPonto   := aParam[2]
    cIdModel   := aParam[3]

    lIsGrid    := ( Len( aParam ) > 3 ) 
    If cIdPonto == 'FORMLINEPOS' 
		nVal  := aParam[4]
		if nVal == 1
			oModelPad  := FWModelActive()
			oModelGrid := oModelPad:GetModel('SDBDETAIL')
			FOR nI := 1 TO oModelGrid:Length()

				cProduto  := oModelGrid:GetValue("DB_PRODUTO",nI)
				cDoc   	  := oModelGrid:GetValue("DB_DOC",nI)
				cSerie 	  := oModelGrid:GetValue("DB_SERIE",nI)
				cCliFor	  := oModelGrid:GetValue("DB_CLIFOR",nI)
				cLoja	  := oModelGrid:GetValue("DB_LOJA",nI)
				cServ	  := oModelGrid:GetValue("DB_SERVIC",nI)
				cTar	  := oModelGrid:GetValue("DB_TAREFA",nI)
				cAtiv	  := oModelGrid:GetValue("DB_ATIVID",nI)

				SDB->(DbSetOrder(7)) //DB_FILIAL+DB_PRODUTO+DB_DOC+DB_SERIE+DB_CLIFOR+DB_LOJA+DB_SERVIC+DB_TAREFA+DB_ATIVID	
				SB1->(DbSetOrder(1))   // FILIAL + CODIGO
				
				If SDB->(DbSeek(xFilial('SDB')+cProduto+cDoc+cSerie+cCliFor+cLoja+cServ+cTar+cAtiv)) .and. SDB->DB_ATIVID == '018' .or. SDB->DB_ATIVID == '010'
					If SB1->(DbSeek(xFilial('SB1')+cProduto)) .and. ALLTRIM(SB1->B1_UM) == "KG"  .and. cTar == '002' .And. .F.

						if nX == 1 
							MsgInfo('Em seguida será necessário informar o peso exato de alguns itens.'+ CRLF +;
							'Antes de alterar o peso do produto certifique que o novo peso será' + CRLF +;
							'atribuido ao item correto, em seguida altere informando o peso exato.' , "ATENÇÃO")
							nX++
						endif 

						nPeso := 0
						while nPeso == 0
							nOpcA := 0
							@ 000,000 TO 250,400 DIALOG oDlg TITLE "ALTERAR PESO DO PRODUTO " + SDB->DB_PRODUTO

							@ 04, 065 SAY "INFORME O PESO CORRETO."

							@ 016, 010 SAY "CODIGO: " + SDB->DB_PRODUTO

							@ 026, 010 SAY "DESCRIÇÃO: " + SB1->B1_DESC

							@ 036, 010 SAY "NR. ITEM: " + SDB->DB_SERIE

							@ 046, 010 SAY "PESO ATUAL: " + cValToChar(SDB->DB_QUANT)

							@ 066, 010 SAY "Informe o Peso!"

							@ 062, 060 GET nPeso Picture PesqPict("SDB","DB_QUANT") SIZE 100,15  PIXEL OF oDlg 

							@ 090, 160 BMPBUTTON TYPE 01 ACTION (nOpcA:=1, GravaPeso(nPeso))
							ACTIVATE DIALOG oDlg CENTER
							if nOpcA <> 1
								MsgAlert("Necessário informar o peso e confirmar", "ATENCÃO")
								nPeso := 0
							endif
						Enddo
					endif
				endif
			Next nI
		endif
	ElseIf cIdPonto == 'FORMCOMMITTTSPOS'
		GravaOrcamento()
	endif
endif

RETURN xRet

Static Function GravaPeso(nPeso)

	Local aArea := GetArea()
	//Local cQuery := ''
//
	//	cQuery += " SELECT DCR_FILIAL,	  "
	//	cQuery += "       DCR.DCR_IDORI,  "
	//	cQuery += "       DCR.DCR_IDDCF,  "
	//	cQuery += "       DCR.DCR_IDMOV,  "
	//	cQuery += "       DCR.DCR_IDOPER, "
	//	cQuery += "       DCR.DCR_SEQUEN  "
	//	cQuery += " FROM  "+RetSqlName("DCR")+" DCR "
	//	cQuery += "       INNER JOIN "+RetSqlName("DCF")+" DCF "
	//	cQuery += "       ON DCF_ID = DCR_IDDCF "
	//	cQuery += " WHERE  DCF_DOCTO = '" + oModelGrid:GetValue("DB_DOC") + "'"
//
	//	cQuery := ChangeQuery(cQuery)
	//	cAliasQry := GetNextAlias()
	//	DbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliasQry,.T.)
			
	If Empty(nPeso)
		msgAlert("Por favor prencher campo com o peso do Produto!")	
	Else
		RecLock('SDB',.F.)
		SDB->DB_QUANT := nPeso
		SDB->DB_QTSEGUM := ConvUM(SDB->DB_PRODUTO,nPeso,0,2) // 2UM
		if SDB->DB_EMPENHO >0
			SDB->DB_EMPENHO := nPeso
			SDB->DB_EMP2 := ConvUM(SDB->DB_PRODUTO,nPeso,0,2)
		endif
		MsUnlock()

		DbSelectArea('DCF')
		DCF->(DbSetOrder(2)) // DCF_FILIAL + DCF_SERVIC + DCF_DOCTO + DCF_SERIE + DCF_CLIFOR + DCF_LOJA + DCF_CODPRO
		If  DCF->(DbSeek(SDB->DB_FILIAL + SDB->DB_SERVIC  + SDB->DB_DOC + SDB->DB_SERIE + SDB->DB_CLIFOR + SDB->DB_LOJA + SDB->DB_PRODUTO))
			RecLock('DCF',.F.)
			DCF->DCF_QUANT := nPeso
			DCF->DCF_QTSEUM := ConvUM(SDB->DB_PRODUTO,nPeso,0,2)
			DCF->(MsUnlock())
		endif

		DbSelectArea('DCR')
		DCR->(DbSetOrder(3))
		
		If DCR->(DbSeek(DCF->DCF_FILIAL+DCF->DCF_ID))
			While DCR->(!Eof()) .and. DCR_IDDCF == SDB->DB_IDDCF
				RecLock('DCR',.F.)
				DCR->DCR_QUANT := nPeso
				DCR->DCR_QTSEUM := ConvUM(SDB->DB_PRODUTO,nPeso,0,2) // 2UM
				MsUnlock()
				DCR->(dbSkip())
			enddo
		endif
			
		DbSelectArea('SC6')
		SC6->(DbSetOrder(1))   // C6_FILIAL + C6_NUM + C6_ITEM + C6_PRODUTO
		If SC6->(DbSeek( SDB->DB_FILIAL + alltrim(SDB->DB_DOC) + alltrim(SDB->DB_SERIE) + SDB->DB_PRODUTO))
			RecLock('SC6',.F.)
			SC6->C6_QTDVEN := nPeso
			SC6->C6_VALOR := (SC6->C6_PRCVEN * nPeso)
			SC6->C6_UNSVEN := ConvUM(SDB->DB_PRODUTO,nPeso,0,2)
			SC6->C6_QTDEMP := nPeso
			SC6->C6_QTDEMP2 := ConvUM(SDB->DB_PRODUTO,nPeso,0,2)
			SC6->(MsUnlock())
		endif
		
		DbSelectArea('SC9')
		SC9->(DbSetOrder(2))  // C9_FILIAL + C9_CLIENTE + C9_LOJA + C9_PEDIDO + C9_ITEM
		If  SC9->(DbSeek(SDB->DB_FILIAL + SDB->DB_CLIFOR+ SDB->DB_LOJA + alltrim(SDB->DB_DOC) + alltrim(SDB->DB_SERIE)))
			while SC9->(!EoF()) .and. SC9->C9_PRODUTO == SDB->DB_PRODUTO
				RecLock('SC9',.F.)
				SC9->C9_QTDLIB := nPeso
				SC9->C9_QTDLIB2 := ConvUM(SDB->DB_PRODUTO,nPeso,0,2)
				SC9->(MsUnlock())
			SC9->(dbSkip())
			enddo
		endif

		SDB->(DbSetOrder(7)) //DB_FILIAL+DB_PRODUTO+DB_DOC+DB_SERIE+DB_CLIFOR+DB_LOJA+DB_SERVIC+DB_TAREFA+DB_ATIVID			
		if SDB->(DBSEEK(xFilial('SDB') + SDB->DB_PRODUTO + SDB->DB_DOC + SDB->DB_SERIE + SDB->DB_CLIFOR + SDB->DB_LOJA + SDB->DB_SERVIC + '003' + '037'))
			RecLock('SDB', .F.)
			SDB->DB_QUANT := nPeso
			SDB->DB_QTSEGUM := ConvUM(SDB->DB_PRODUTO,nPeso,0,2) // 2UM
			if SDB->DB_EMPENHO >0
				SDB->DB_EMPENHO := nPeso
				SDB->DB_EMP2 := ConvUM(SDB->DB_PRODUTO,nPeso,0,2)
			endif
			SDB->(MsUnlock())
		endif
	endif
	//(cAliasQry)->( dbCloseArea() )
	RestArea(aArea)
	Close(oDlg)
Return 

Static Function GravaOrcamento()
	Local cQry, cTab, aBkp, aLinha, cNumOrc, nSaveSX8
	Local aArrSL1 := {}
	Local aArrSL2 := {}
	Local aArrSL4 := {}
	Local aEstorn := {}
	Local cPedido := Trim(SDB->DB_DOC)
	Local cTabPad := PADR("1",TamSX3("LR_TABELA")[1])
	Local nTotal  := 0
	Local lRet    := .F.
	
	// Expedição + Conferência + Conferência
	If SDB->DB_FILIAL == '15' .And. SDB->DB_SERVIC == "001" .And. SDB->DB_TAREFA == "003" .And. SDB->DB_ATIVID == "037" .And. SDB->DB_STATUS $ "MA" .And. !Empty(cPedido)
		
		SC5->(dbSetOrder(1))//C5_FILIAL+C5_NUM
		If SC5->(dbSeek(SDB->DB_FILIAL+cPedido)) .And. SC5->C5_TABELA == '011'

			aBkp := GetArea()
			cTab := GetNextAlias()
			
			cQry := "SELECT SC6.C6_NUM, SC6.C6_ITEM, SC6.C6_PRODUTO, ISNULL(SC9.C9_QTDLIB,0) AS C9_QTDLIB, SC9.C9_LOCAL, SC9.C9_PRCVEN,"
			cQry += " SC9.C9_LOTECTL, SC9.C9_ENDPAD, ISNULL(SC9.R_E_C_N_O_,0) AS C9_RECNO"
			cQry += " FROM " + RetSQLName("SC6") + " SC6"
			cQry += " LEFT OUTER JOIN " + RetSQLName("SC9") + " SC9 ON SC9.D_E_L_E_T_ = ' '"
			cQry += " AND SC9.C9_FILIAL = SC6.C6_FILIAL"
			cQry += " AND SC9.C9_PEDIDO = SC6.C6_NUM"
			cQry += " AND SC9.C9_ITEM = SC6.C6_ITEM"
			cQry += " AND SC9.C9_NFISCAL = ' '"
			cQry += " AND SC9.C9_BLWMS = '05'"
			cQry += " AND SC9.C9_LOTECTL <> ' '"
			cQry += " AND SC9.C9_ENDPAD <> ' '"
			cQry += " WHERE SC6.D_E_L_E_T_ = ' '"
			cQry += " AND SC6.C6_FILIAL = '"+SC6->(XFILIAL("SC6"))+"'"
			cQry += " AND SC6.C6_BLOQUEI = ' '"
			cQry += " AND SC6.C6_QTDVEN > SC6.C6_QTDENT"
			cQry += " AND SC6.C6_NUM = '"+cPedido+"'"
			cQry += " ORDER BY SC6.C6_NUM, SC6.C6_ITEM"
			
			dbUseArea(.T.,"TOPCONN",TCGenQry(,,ChangeQuery(cQry)),cTab,.F.,.F.)
			While !Eof()
				
				If Empty((cTab)->C9_QTDLIB)   // Caso tenha item sem liberação
					aArrSL2 := {}
					Exit
				Endif
				
				// Posiciona no Cadastro do Produto
				SB1->(dbSetOrder(1))
				SB1->(dbSeek(XFILIAL("SB1")+(cTab)->C6_PRODUTO))
				
				aLinha := {}
				AAdd( aLinha , {"LR_PRODUTO", (cTab)->C6_PRODUTO , NIL} )
				AAdd( aLinha , {"LR_DESCRI" , SB1->B1_DESC       , NIL} )
				AAdd( aLinha , {"LR_LOCAL"  , (cTab)->C9_LOCAL   , NIL} )
				AAdd( aLinha , {"LR_TABELA" , cTabPad            , NIL} )
				AAdd( aLinha , {"LR_UM"     , SB1->B1_UM         , NIL} )
				//AAdd( aLinha , {"LR_DESCPRO", 0                  , NIL} )
				//AAdd( aLinha , {"LR_VEND"   , cVenVenda             , NIL} )
				AAdd( aLinha , {"LR_QUANT"  , (cTab)->C9_QTDLIB  , NIL} )
				AAdd( aLinha , {"LR_VRUNIT" , (cTab)->C9_PRCVEN  , NIL} )
				AAdd( aLinha , {"LR_PRCTAB" , (cTab)->C9_PRCVEN  , NIL} )
				//AAdd( aLinha , {"LR_DESC"   , 0                  , NIL} )
				//AAdd( aLinha , {"LR_VALDESC", 0                  , NIL} )
				AAdd( aLinha , {"LR_VLRITEM", (cTab)->C9_QTDLIB * (cTab)->C9_PRCVEN, NIL} )
				AAdd( aLinha , {"LR_LOTECTL", (cTab)->C9_LOTECTL , NIL} )
				AAdd( aLinha , {"LR_LOCALIZ", (cTab)->C9_ENDPAD  , NIL} )
				AAdd( aArrSL2 , aClone(aLinha) )
				
				nTotal += (cTab)->C9_QTDLIB * (cTab)->C9_PRCVEN
				
				AAdd( aEstorn , (cTab)->C9_RECNO )    // Adiciona para estorno da liberação
				
				dbSkip()
			Enddo
			dbCloseArea()
			
			If !Empty(aArrSL2)
				nSaveSX8 := GetSx8Len()                 // Variavel que controla numeracao
				cNumOrc  := ProxOrcamento()
				
				BeginTran()
				
				FWMsgRun(Nil, {|oSay| EstornaPedido(aEstorn)          }, "Aguarde...", "Estornando Liberação")
				FWMsgRun(Nil, {|oSay| lRet := EliminaResiduo(cPedido) }, "Aguarde...", "Eliminando Resíduo")
				
				If lRet
					aAdd( aArrSL4, {} )
					aAdd( aTail(aArrSL4), {"L4_DATA"    , dDataBase    , NIL} )
					aAdd( aTail(aArrSL4), {"L4_VALOR"   , nTotal       , NIL} )
					aAdd( aTail(aArrSL4), {"L4_FORMA"   , "R$"         , NIL} )
					aAdd( aTail(aArrSL4), {"L4_ADMINIS" , " "          , NIL} )
					aAdd( aTail(aArrSL4), {"L4_NUMCART" , " "          , NIL} )
					aAdd( aTail(aArrSL4), {"L4_FORMAID" , " "          , NIL} )
					aAdd( aTail(aArrSL4), {"L4_MOEDA"   , 0            , NIL} )
					
					SC5->(dbSetOrder(1))
					SC5->(dbSeek(XFILIAL("SC5")+cPedido))
					SA1->(dbSetOrder(1))
					SA1->(dbSeek(XFILIAL("SA1")+SC5->C5_CLIENTE+SC5->C5_LOJACLI))
					SA3->(dbSetOrder(1))
					SA3->(dbSeek(XFILIAL("SA3")+SC5->C5_VEND1))
					
					AAdd( aArrSL1 , { "LQ_NUM"    , cNumOrc       , NIL} )
					AAdd( aArrSL1 , { "LQ_VEND"   , SA3->A3_COD   , NIL} )
					AAdd( aArrSL1 , { "LQ_COMIS"  , 0             , NIL} )
					AAdd( aArrSL1 , { "LQ_CLIENTE", SA1->A1_COD   , NIL} )
					AAdd( aArrSL1 , { "LQ_LOJA"   , SA1->A1_LOJA  , NIL} )
					AAdd( aArrSL1 , { "LQ_NOMCLI" , SA1->A1_NOME  , NIL} )
					AAdd( aArrSL1 , { "LQ_TIPOCLI", SA1->A1_TIPO  , NIL} )
					AAdd( aArrSL1 , { "LQ_DESCONT", 0             , NIL} )
					AAdd( aArrSL1 , { "LQ_DTLIM"  , dDataBase     , NIL} )
					AAdd( aArrSL1 , { "LQ_EMISSAO", dDataBase     , NIL} )
					AAdd( aArrSL1 , { "LQ_CONDPG" , "001"         , NIL} )
					AAdd( aArrSL1 , { "LQ_NUMMOV" , "1 "          , NIL} )
					AAdd( aArrSL1 , { "LQ_PARCELA", Len(aArrSL4)  , NIL} )
					AAdd( aArrSL1 , { "LQ_FORMPG" , "R$"          , NIL} )
					AAdd( aArrSL1 , { "LQ_DINHEIR", nTotal        , NIL} )
					AAdd( aArrSL1 , { "LQ_ENTRADA", nTotal        , NIL} )
					AAdd( aArrSL1 , { "LQ_ENTREGA", dDataBase     , NIL} )
					
					FWMsgRun(Nil, {|oSay| lRet := CriaOrcamento(aArrSL1,aArrSL2,aArrSL4) }, "Aguarde...", "Gravando Orçamento")
				Endif
				
				If lRet
					EndTran()
					
					// Confirma SX8
					While ( GetSx8Len() > nSaveSX8 )
						ConfirmSX8()
					Enddo
					
					FWAlertSuccess("Orçamento "+cNumOrc+" gravado com sucesso !")
				Else
					DisarmTransaction()
					
					// Cancela SX8
					While ( GetSx8Len() > nSaveSX8 )
						RollBackSX8()
					Enddo
				Endif
			Endif
			
			RestArea(aBkp)
		EndIf
	Endif

Return lRet

Static Function CriaOrcamento(aArrSL1,aArrSL2,aArrSL4)
	Local lModAux := nModulo  // Restaura o módulo acessado
	Local lBkpInc := INCLUI
	Local lBkpAlt := ALTERA
	Local lRet    := .F.
	
	SetFunName("LOJA701")
	
	lMsHelpAuto := .T.
	lMsErroAuto := .F.
	
	nModulo     := 12       // Define como módulo do Loja
	lFiscal     := .F.
	INCLUI      := .T.
	ALTERA      := .F.
	lAutomatoX  := .F.
	
	MSExecAuto({|a,b,c,d,e,f,g,h| Loja701(a,b,c,d,e,f,g,h)},.F.,3,"","",{},aArrSL1,aArrSL2,aArrSL4)
	
	If lMsErroAuto
		MostraErro()
	Else
		lRet := .T.
	Endif
	
	nModulo := lModAux  // Restaura o módulo acessado
	INCLUI  := lBkpInc
	ALTERA  := lBkpAlt
	
	SetFunName("WMSA331")

Return lRet

Static Function EstornaPedido(aEstorn)
	Local nX
	
	For nX:=1 To Len(aEstorn)
		SC9->(dbGoTo(aEstorn[nX]))    // Posiciona no regisro a ser estornado
		SC9->(A460Estorna())          // Estorna a liberação
	Next

Return

Static Function EliminaResiduo(cPedido)
	Local nX
	Local aReg  := {}
	Local cSeek := SC5->(XFILIAL("SC5"))+cPedido
	Local lRet  := .F.
	
	SC5->(dbSetOrder(1))
	SC5->(dbSeek(cSeek))
	SC6->(dbSetOrder(1))
	SC6->(dbSeek(cSeek,.T.))
	
	While !SC6->(Eof()) .And. SC6->C6_FILIAL+SC6->C6_NUM == cSeek
		
		If Empty(SC6->C6_RESERVA) .And. !SC6->C6_BLQ$"R #S "
			aAdd(aReg, SC6->(RecNo()))
		EndIf
		
		SC6->(dbSkip())
	Enddo
	
	If !Empty(aReg)
		// Elimina residuo de itens e atualiza cabecalho do pedido.
		For nX := 1 To Len(aReg)
			SC6->(dbGoTo(aReg[nX]))				 
			If !(lRet := MaResDoFat(nil, .T., .F., /*@nVlrDep*/))
				Exit
			Endif
		Next nX
		
		If lRet
			MaLiberOk({ SC5->C5_NUM }, .T.)
		Endif
	Else
		lRet := .T.
	Endif

Return lRet

Static Function ProxOrcamento()
	Local aArea := SL1->(GetArea())
	Local cNum  := GetSxENum("SL1","L1_NUM")
	Local cMay  := Alltrim(xFilial("SL1"))+cNum
	Local nTent := 0
	
	FreeUsedCode()
	SL1->(DbSetOrder(1))
	
	//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
	//³ Se dois orcamentos iniciam ao mesmo tempo a MayIUseCode impede que ambos utilizem o mesmo numero.³
	//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
	While SL1->(MsSeek(xFilial("SL1")+cNum)) .OR. !MayIUseCode(cMay)
		If ++nTent > 20
			Final("Impossivel gerar numero sequencial de orcamento correto. Informe ao administrador do sistema.")
		EndIf
		While (GetSX8Len() > nSaveSx8)
			ConfirmSx8()
		Enddo
		cNum := GetSxENum("SL1","L1_NUM")
		FreeUsedCode()
		cMay := Alltrim(xFilial("SL1"))+cNum
	Enddo
	SL1->(RestArea(aArea))

Return cNum
