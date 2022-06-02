#INCLUDE "Font.ch"
#Include "Protheus.ch" 
#include "rwmake.ch"                                            
#INCLUDE "fileio.ch"

/*_______________________________________________________________________________
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¦¦+-----------+------------+-------+----------------------+------+------------+¦¦
¦¦¦ Função    ¦ QBG007     ¦ Autor ¦    Edson Sales       ¦ Data ¦ 			  ¦¦¦
¦¦+-----------+------------+-------+----------------------+------+------------+¦¦
¦¦¦ Descriçäo ¦ Fonte  para Atualizar SD1 com descição dos produtos            ¦¦
¦¦+-----------+---------------------------------------------------------------+¦¦
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯*/
User Function QBG007()

	Local cPerg := "QBG007"
	Private oDlg,oButConf,oButFechar,oButPar,oSay1,oSay2//oButArq

	Pergunte(cPerg,.F.)
	DEFINE DIALOG oDlg TITLE "Atualizar Descrição"  FROM 240,1 TO 440,440 PIXEL
	@ 10,10 TO 90,210
	oSay1  := TSay():New(20,018,{||'Este programa tem como objetivo atualizar descrição'},oDlg,,,,,,.T.,CLR_BLACK,CLR_WHITE,200,12)
	oSay2  := TSay():New(28,018,{||'conforme os parametros especificados ou arquivo selecionado pelo usuário'},oDlg,,,,,,.T.,CLR_BLACK,CLR_WHITE,220,12)
	oButPar		:= TButton():New(70,060, "Parametros",oDlg,{|| Pergunte(cPerg,.T.)  },40,10,,,.F.,.T.,.F.,,.F.,,,.F.)
	oButArq		:= TButton():New(70,110, "Confirmar" ,oDlg,{|| Processa({|| AtuaDesc(),"Atualização de dados na D1 - XML",.F.}) },40,10,,,.F.,.T.,.F.,,.F.,,,.F.)
	oButFechar	:= TButton():New(70,160, "Fechar"    ,oDlg,{||  oDlg:End()  },40,10,,,.F.,.T.,.F.,,.F.,,,.F.)
	//oButConf	:= TButton():New(70,170, "Arquivos"  ,oDlg,{|| fSelArq()  },40,10,,,.F.,.T.,.F.,,.F.,,,.F.)

	Activate Dialog  oDlg Centered

Return 

Static Function AtuaDesc()

	Local 	 cFil		:= mv_par01
	Local  	 dDatade	:= mv_par02
	Local	 dDataate	:= mv_par03
	Local    cDest 		:= '\xml\'
	Local 	 oReport 	:= nil
	Local	 cQuery		:= ''
	Private  aRel  := {}

	Private  cArquivo	:= alltrim(mv_par04)
	Private  cChvNfe 	:= Substr(cArquivo,Rat("\",cArquivo )+1,44)
	Private  nAbre

	If Empty(cArquivo)
		If Empty(dDatade) .or. Empty(dDataate) 		// validação de data
			Aviso("Problema ao abrir arquivo", "Por favor, informa data.", {"Fechar"}, 1)
			return
		EndIf
	
		If Empty(cFil)  							// validação da filial
			Aviso("Problema ao abrir arquivo", "Por favor, informe a Filial.", {"Fechar"}, 1)
			Return
		EndIf
	EndIF
	

	SF1->(DbSetOrder(8))	//F1_FILIAL+F1_CHVNFE
	SD1->(DbSetOrder(1))	//D1_FILIAL+D1_DOC+D1_SERIE+D1_FORNECE+D1_LOJA+D1_COD+D1_ITEM

	
	cQuery += " SELECT DISTINCT SF1.F1_CHVNFE CHAVE, "
	cQuery += "                 SF1.F1_DOC DOC, "
	cQuery += "                 SF1.F1_SERIE SERIE, "
	cQuery += "                 SF1.F1_FORNECE FORNECE, "
	cQuery += "                 SF1.F1_LOJA LOJA, "
	cQuery += "                 IIF(SF1.F1_TIPO='D',SA1.A1_NOME,SA2.A2_NOME) NOME,  "
	cQuery += "                 SF1.R_E_C_N_O_ RECSF1, "
	cQuery += "                 (SELECT COUNT(*) FROM SD1010 X WHERE X.D_E_L_E_T_ ='' AND X.D1_FILIAL=SF1.F1_FILIAL AND X.D1_DOC = SF1.F1_DOC AND X.D1_SERIE = SF1.F1_SERIE AND X.D1_FORNECE=SF1.F1_FORNECE AND X.D1_LOJA=SF1.F1_LOJA ) AS ITEM "
	cQuery += " FROM  " + RetSqlName("SF1") + "  SF1 "
	cQuery += " INNER  JOIN  " + RetSqlName("SD1") + " SD1 ON SD1.D_E_L_E_T_ = '' AND SD1.D1_FILIAL = SF1.F1_FILIAL AND SD1.D1_DOC = SF1.F1_DOC AND SD1.D1_SERIE = SF1.F1_SERIE AND SD1.D1_FORNECE=SF1.F1_FORNECE AND SD1.D1_LOJA=SF1.F1_LOJA AND SD1.D1_TES!='' "
	cQuery += " LEFT OUTER JOIN  " + RetSqlName("SA2") + " SA2 ON SA2.D_E_L_E_T_ = '' AND SA2.A2_COD = SF1.F1_FORNECE AND SA2.A2_LOJA=SF1.F1_LOJA "
	cQuery += " LEFT OUTER JOIN SA1010 SA1 ON SA1.D_E_L_E_T_ = '' AND SA1.A1_COD = SF1.F1_FORNECE AND SA1.A1_LOJA=SF1.F1_LOJA  "
	cQuery += " LEFT OUTER JOIN " + RetSqlName("SZX") + " SZX ON SZX.D_E_L_E_T_ = '' AND SZX.ZX_FILIAL=SD1.D1_FILIAL AND SZX.ZX_DOC=SD1.D1_DOC AND SZX.ZX_SERIE=SD1.D1_SERIE AND SZX.ZX_FORNECE=SD1.D1_FORNECE AND SZX.ZX_LOJA=SD1.D1_LOJA AND SZX.ZX_ITEM=SD1.D1_ITEM "
	cQuery += " WHERE   SF1.D_E_L_E_T_='' "
	cQuery += "         AND SF1.F1_FILIAL 	 = '"  + cFil   + "'"
	cQuery += "         AND SF1.F1_DTDIGIT BETWEEN '" +DTOS(dDatade) +"' AND '" + DTOS(dDataate)+"'"
	cQuery += "         AND SF1.F1_ESPECIE != 'CTE' "
	cQuery += "         AND SF1.F1_CHVNFE <> '' "
	cQuery += "         AND SF1.F1_TIPO IN ('N','D') "
	cQuery += "         AND SF1.F1_FORMUL = '' "
	cQuery += "         AND SF1.F1_FORNECE != '000010' "
	cQuery += "         AND SZX.ZX_FILIAL IS NULL "


	//cQuery += " ORDER  BY SD1.D1_ITEM "
	
	if !Empty(cChvNfe)							 // se o arquivo for Selecionado, chama a função para atualizar a tebela.   
		If ":" $ cArquivo 						 //Checa se não está na pasta do sistema (se não está, então contém ":")
			If File(cDest+cChvNfe+'.xml')
				FErase(cDest+cChvNfe+'.xml') 	 // Se existir, apaga
			Endif
			CpyT2S(cArquivo , cDest)  			 // copia para o servidor
			cArquivo := cDest + cChvNfe + '.xml'   
		Endif
		cQuery += " AND SF1.F1_CHVNFE = '" + cChvNfe + "'"
	Endif

	cQuery := ChangeQuery(cQuery)
	dbUseArea(.T.,"TOPCONN",tcGenQry(,,cQuery),"TMP",.F.,.F.)
	DbSelectArea('TMP')

	while !TMP->(EoF()) 				 // se o arquivo nao for selecionado ,chama a função para atualizar a tebela e usar a busca da query.
		SF1->(dbGoTo(TMP->RECSF1))
		cArquivo := (cDest + TMP->CHAVE + ".xml")
		//If type("cArquivo") == 'U'
		//	alert("deu merda")
		//EndIf
		nAbre := fOpen(cArquivo,0)
		If nAbre != -1
			AtualTab(cArquivo)
		else //Monta array SE NÃO ABRIR O DOC
			AADD(aRel,{TMP->DOC,RIGHT('000'+TMP->SERIE,3),TMP->FORNECE,TMP->NOME,TMP->CHAVE, 'XML nao encontrado'})	
		EndIf
		TMP->(DbSkip())
	Enddo
		
	TMP->(DBCLOSEAREA())
	
	If !Empty(aRel)
		MsgInfo("Importação concluida, Segue Relatorio Documentos não atualizados ","Informação")
		oReport := DocRel(aRel)
		oReport:PrintDialog()	
	Else
		MsgInfo("Importação concluida com Sucesso","Sucesso.")
	EndIf
RETURN

Static Function AtualTab(cArquivo)
	Local cWarning 	:= ""
	Local cError   	:= ""
	Local aProduto  :={}
	Local oXML
	Local oDet
	LOCAL nX,nFator,cTipo

	SD1->(DbSetOrder(1)) //D1_FILIAL+D1_DOC+D1_SERIE+D1_FORNECE+D1_LOJA+D1_COD+D1_ITEM 
	SZX->(DbSetOrder(1)) //ZX_FILIAL+ZX_DOC+ZX_SERIE+ZX_FORNECE+ZX_LOJA+ZX_ITEM 

	oXML := XmlParserFile( cArquivo , "_", @cError, @cWarning ) 	//Efetua abertura do arquivo XML
	oDet := oXML:_nfeproc:_nfe:_infnfe:_det 	 					//Detalhes da nota (produtos, etc)

	If !Empty(cError) 							 					//Caso tenha ocorrido erro:
		cMsgErr := cError
		Alert(cMsgErr)
	Else
		If Valtype(oDet) == 'O' 
			If	SD1->(dbSeek(SF1->(F1_FILIAL+F1_DOC+F1_SERIE+F1_FORNECE+F1_LOJA)))
				nQtdXml :=Val(oDet:_prod:_qcom:text)  // traz a quant do xml)
				// calcular o fator de conversao
				if SD1->D1_QUANT > nQtdXml
					nFator := SD1->D1_QUANT/nQtdXml
					cTipo := "D" 
				else
					nFator := (nQtdXml/SD1->D1_QUANT)
					cTipo := "M"
				endIf

				lInclui := !SZX->(dbSeek(SF1->(F1_FILIAL+F1_DOC+F1_SERIE+F1_FORNECE+F1_LOJA)))

				Reclock ("SZX",lInclui)
					SZX->ZX_FILIAL 	:= SD1->D1_FILIAL
					SZX->ZX_DOC 	:= SD1->D1_DOC
					SZX->ZX_SERIE 	:= SD1->D1_SERIE
					SZX->ZX_FORNECE := SD1->D1_FORNECE
					SZX->ZX_LOJA	:= SD1->D1_LOJA
					SZX->ZX_ITEM	:= SD1->D1_ITEM
					SZX->ZX_DESCXML := oDet:_prod:_xprod:text
					SZX->ZX_UMXML	:= oDet:_prod:_ucom:text
					SZX->ZX_TPCONV 	:= cTipo
					SZX->ZX_CONVXML := nFator
				SZX->(MsUnlock ())

				If TMP->ITEM > 1
					AADD(aRel,{TMP->DOC,RIGHT('000'+TMP->SERIE,3),TMP->FORNECE,TMP->NOME,TMP->CHAVE, 'Itens divergentes XML = 1, Sistema = ' +AllTrim(str(TMP->ITEM))})
				Endif

			Endif	
		else
			If SD1->(dbSeek(SF1->(F1_FILIAL+F1_DOC+F1_SERIE+F1_FORNECE+F1_LOJA)))
				while !SD1->(EOF()) .and. SF1->(F1_FILIAL+F1_DOC+F1_SERIE+F1_FORNECE+F1_LOJA) == SD1->(D1_FILIAL+D1_DOC+D1_SERIE+D1_FORNECE+D1_LOJA)
					aAdd(aProduto,{SD1->D1_COD,SD1->D1_ITEM})
					SD1->(dbskip())
				Enddo
			EndIf
			
			ASORT(aProduto, , , { | x,y | x[2] < y[2] } )
			
			for nX := 1 To len(oDet)

				If SD1->(dbSeek(SF1->(F1_FILIAL+F1_DOC+F1_SERIE+F1_FORNECE+F1_LOJA)+aProduto[nx,1]+aProduto[nx,2]))

					nQtdXml :=Val(oDet[nX]:_prod:_qcom:text)  // traz a quant do xml)
					// calcular o fator de conversao
					if SD1->D1_QUANT > nQtdXml
						nFator := SD1->D1_QUANT/nQtdXml
						cTipo := "D" 
					else
						nFator := (nQtdXml/SD1->D1_QUANT)
						cTipo := "M"
					endIf

					lInclui := !SZX->(dbSeek(SF1->(F1_FILIAL+F1_DOC+F1_SERIE+F1_FORNECE+F1_LOJA)+SD1->D1_ITEM))

					Reclock("SZX",lInclui)
					SZX->ZX_FILIAL 	:= SD1->D1_FILIAL
					SZX->ZX_DOC 	:= SD1->D1_DOC
					SZX->ZX_SERIE 	:= SD1->D1_SERIE
					SZX->ZX_FORNECE := SD1->D1_FORNECE
					SZX->ZX_LOJA	:= SD1->D1_LOJA
					SZX->ZX_ITEM	:= SD1->D1_ITEM
					SZX->ZX_DESCXML := oDet[nx]:_prod:_xprod:text
					SZX->ZX_UMXML	:= oDet[nX]:_prod:_ucom:text
					SZX->ZX_TPCONV 	:= cTipo
					SZX->ZX_CONVXML := nFator 
					SZX->(MsUnlock ())
					
				EndIf
			Next nX
			If !TMP->ITEM == LEN(oDet)
				AADD(aRel,{TMP->DOC,RIGHT('000'+TMP->SERIE,3),TMP->FORNECE,TMP->NOME,TMP->CHAVE, 'Itens divergentes XML =' + AllTrim(STR(LEN(oDet)))  + ', Sistema = ' +AllTrim(str(TMP->ITEM))})
			Endif
		Endif
	EndIf

	fClose(cArquivo)
Return

Return

Static Function DocRel(aRel)
	Local oReport := Nil
	Local oSection1:= Nil

	//Local aOrdem := {'Doc/Emissao'}

	oReport := TReport():New("Relatorio_rel001","Relatório de Documentos não atualizados",,{|oReport| ReportPrint(oReport,aRel)},"Relatório de Documentos que não atualizaram a descrição")
	oReport:SetLandscape() // Seta como paisagem
	oReport:SetTotalInLine(.F.)
	oReport:SetLandscape()
	oReport:nFontBody	:= 8
	oReport:nLineHeight	:= 48

	//Relatorio Sintetico
	oSection1:= TRSection():New(oReport, "Analise", {},, .F., .T.)
	TRCell():New(oSection1,"DOC"			,"TRB","Documento"	 	 ,"@!"			,9)
	TRCell():New(oSection1,"SERIE"			,"TRB","Serie",           "@!"			,3,)
	TRCell():New(oSection1,"FORNEC"			,"TRB","N.Fornecedor"	,"@!"	    	,6,)
	TRCell():New(oSection1,"NOME"			,"TRB","Fornecedor"		 ,"@!"	    	,30,)
	TRCell():New(oSection1,"CHAVE"			,"TRB","Chave"	    	 ,"@!"			,44,)
	TRCell():New(oSection1,"MENSAGEM"		,"TRB","Mensagem"	     ,""			,45,)
	
Return(oReport)

Static Function ReportPrint(oReport,aRel)
	Local nI := 0
	Local oSection1 := oReport:Section(1)
	//inicializo a de acordo com a ordem escolhida
	oReport:IncMeter()
	oSection1:Init()
	//Set os valores no objeto
	for nI := 1 to len(aRel)
		oSection1:Cell(	"DOC"  		):SetValue(aRel[ni, 1] )
     	oSection1:Cell(	"SERIE"		):SetValue(aRel[ni, 2] )
		oSection1:Cell(	"FORNEC"	):SetValue(aRel[ni, 3] )
		oSection1:Cell(	"NOME"  	):SetValue(aRel[ni, 4] )		
		oSection1:Cell(	"CHAVE"  	):SetValue(aRel[ni, 5] )
     	oSection1:Cell(	"MENSAGEM"  ):SetValue(aRel[ni, 6] )

		oSection1:PrintLine()
		oReport:IncMeter()
	next nI
	oSection1:Finish()
	oReport:ThinLine()
Return
