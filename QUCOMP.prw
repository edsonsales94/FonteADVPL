#Include "Protheus.ch" 
#include "rwmake.ch"                                            
#Include "TbiConn.ch"
#Include 'TOTVS.ch'
/*_______________________________________________________________________________
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¦¦+-----------+------------+-------+----------------------+------+------------+¦¦
¦¦¦ Programa  ¦ QUCOMP03    ¦ Autor ¦ Edson Sales         ¦ Data ¦ 06/01/2022 ¦¦¦
¦¦+-----------+------------+-------+----------------------+------+------------+¦¦
¦¦¦ Descriçäo ¦ Função para inserir data de apresentação 'CSV' na F1_XDTAPRE  ¦¦¦ 
¦¦+-----------+---------------------------------------------------------------+¦¦
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯*/

#DEFINE NOME_REL 'Relatorio_Doc_nao_encontrados'

User Function QUCOMP03()

	Local oTela

	@ 200,1 TO 380,380 DIALOG  oTela TITLE OemToAnsi("IMPORTAR CSV/GRAVAR DATA")
	@ 02,10 TO 080,190
	@ 10,018 Say "Este programa tem como objetivo importar arquivo CSV,"
	@ 18,018 Say "para inserção da Data de Apresentação na base de dados."

	@ 70,128 BMPBUTTON TYPE 01 ACTION ValidPerg()
	@ 70,158 BMPBUTTON TYPE 02 ACTION Close( oTela)   
	//@ 70,100 BMPBUTTON TYPE 05 ACTION Pergunte(cPerg,.T.)

	Activate Dialog  oTela Centered

Return

Static Function ValidPerg()
	Local cPerg := "QUDTAPRE"
	if Pergunte(cPerg, .T.)
		LerCsv()
	EndIf
	
Return

Static Function LerCsv()

    Local   cArq := ALLTRIM(mv_par01)
	Local 	xLinha
	//Local lPrim := .T.
	Local aCampos := {}
	Local aDados  := {}
	Local aItem	  := {}
	Local nA 	  :=  0
	Local nX 
	Local oReport := nil
	Local cDoc,cSerie,cCnpj
	Private aRel  := {}
	Private aRRaa := {}

	
	FT_FUSE(cArq)  // Abrir Arquivo
	ProcRegua(FT_FLASTREC())
	FT_FGOTOP()

		While !FT_FEOF()
			IncProc("Lendo arquivo texto...")
			xLinha := FT_FREADLN()

			If nA < 1
				aCampos := Strtokarr(xLinha,'";"')
				nA++
			Else
				aItem := Strtokarr(xLinha,'";"')
				if Empty(aDados) .Or. aScan(aDados,{|x| x[2] == aItem[2]}) == 0
					AADD(aDados,aItem)
				EndIf
			EndIf
		FT_FSKIP()
		EndDo

	SF1->(DbSetOrder(8)) //F1_FILIAL+F1_CHVNFE
	SA2->(DbSetOrder(3))//A2_FILIAL+A2_CGC	
	
	for nX := 1 to len(aDados)
		cDoc	:= SUBSTR(aDados[nX][2],26,9) // Numero Documento 
		cSerie	:= SUBSTR(aDados[nX][2],23,3) // Numero de Serie
		cCnpj	:= SUBSTR(aDados[nX][2],7,14) // CNPJ Fornecedor

		If  cSerie == '000' .or. Left(cSerie,2) == '00' // Se numero de serie tiver 2 zeros na frente ou for composta por 3 zeros
			cSerie := SubStr(cSerie,-1) // Sustring para remover os zero
		elseif Left(cSerie,2) != '00' .and. Left(cSerie,1) == '0' //Se numero de serie tiver 1 zero na frente
			cSerie := SubStr(cSerie,-2)	// Sustring para remover o zero		
		endif
			 //percorre itens posiciona na Chave
			If SF1->(DBSEEK(xFilial('SF1')+aDados[nX][2])) .and. SF1->F1_DOC == Right('0000000'+ALLTRIM(CValToChar(Val(aDados[nx][3]))),9)
				RecLock('SF1',.F.)
					SF1->F1_XDTAPRE := CTOD(aTail(aDados[nx])) 	//ultima possição do array / Data de apresentação
				MsUnlock()
			//posicionar no fornecedor campo A2_CGC 
			Elseif SA2->(DBSEEK(xFilial('SA2')+cCnpj))
				SF1->(DbSetOrder(1))//F1_FILIAL+F1_DOC+F1_SERIE+F1_FORNECE+F1_LOJA+F1_TIPO
				//pecorrer itens posicinado no documento e serie fornecedor etc.
				if SF1->(DBSEEK(xFilial('SF1')+cDoc+padr(cSerie,3)+SA2->A2_COD+SA2->A2_LOJA))
					RecLock('SF1',.F.)
						SF1->F1_XDTAPRE := CTOD(aTail(aDados[nx])) 	//ultima possição do array / Data de apresentação
					MsUnlock()
				endif
			Else  	// Se não encontrar o arquivo
				AADD(aRel,aDados[nx])	//Monta array
			EndIF
	next nX
	fClose(cArq) 
	If !Empty(aRel)
		oReport := DocRel(aRel)
		oReport:PrintDialog()
	Else
		MsgInfo("Importação concluida com Sucesso!!!","Sucesso.")
	EndIf
Return 

/*_______________________________________________________________________________
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¦¦+-----------+------------+-------+----------------------+------+------------+¦¦
¦¦¦ Programa  ¦ QUDTAPRE    ¦ Autor ¦ Edson Sales         ¦ Data ¦ 06/01/2022 ¦¦¦
¦¦+-----------+------------+-------+----------------------+------+------------+¦¦
¦¦¦ Descriçäo ¦       Imprimir Relatorio de documentos não encontrados        ¦¦¦ 
¦¦+-----------+---------------------------------------------------------------+¦¦
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯*/

Static Function DocRel(aRel)
	Local oReport := Nil
	Local oSection1:= Nil

	//Local aOrdem := {'Doc/Emissao'}

	oReport := TReport():New(NOME_REL,"Relatório de Documentos Não Encontrados",,{|oReport| ReportPrint(oReport,aRel)},"Relatório de Documentos Não Encontrados")
	oReport:SetLandscape() // Seta como paisagem
	oReport:SetTotalInLine(.F.)
	oReport:nFontBody	:= 8
	oReport:nLineHeight	:= 48

	//Relatorio Sintetico
	oSection1:= TRSection():New(oReport, "Analise", {},, .F., .T.)
	TRCell():New(oSection1,"DOC"			,"TRB","Documento"	 	 ,"@!"			,09,)
	TRCell():New(oSection1,"EMISSAO"		,"TRB","Dt. Emissao"	 ,"@!"			,10,)
	TRCell():New(oSection1,"CODG CFOP"		,"TRB","Cod CFOP"		 ,"@!"	    	,04,)
	TRCell():New(oSection1,"CHAVE"			,"TRB","Chave"	    	 ,"@!"			,44,)
	TRCell():New(oSection1,"FORNEC"			,"TRB","Fornecedor"		 ,"@!"	    	,65,)
	TRCell():New(oSection1,"DT APRES"		,"TRB","Dt. Apresentaçao","@!"			,10,)

Return(oReport)

Static Function ReportPrint(oReport,aRel)
	Local nI := 0
	Local oSection1 := oReport:Section(1)
	//inicializo a de acordo com a ordem escolhida
	oReport:IncMeter()
	oSection1:Init()
	//Set os valores no objeto
	for nI := 1 to len(aRel)
		oSection1:Cell(	"DOC"  		):SetValue(aRel[ni, 3] )
     	oSection1:Cell(	"EMISSAO"	):SetValue(aRel[ni, 4] )	
		oSection1:Cell(	"CODG CFOP"	):SetValue(aRel[ni, 9] )
		oSection1:Cell(	"CHAVE"  	):SetValue(aRel[ni, 2] )
     	oSection1:Cell(	"FORNEC"	):SetValue(aRel[ni,10] ) 	
		oSection1:Cell(	"DT APRES"	):SetValue(aTail(aRel[ni]))
     	
		oSection1:PrintLine()
		oReport:IncMeter()
	next nI
	oSection1:Finish()
	oReport:ThinLine()

Return
