#Include "Protheus.ch" 
#include "rwmake.ch"                                            
#Include "TbiConn.ch"
#Include 'TOTVS.ch'
#Include "Protheus.ch"
#Include "TopConn.ch"

/*/
+-------------------------------------------------------------------------------+
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¦¦+-----------+------------+-------+----------------------+------+------------+¦¦
¦¦¦ Programa  ¦ CTBGASTO   ¦ Autor ¦Edson Sales           ¦ Data ¦ 30/12/2022 ¦¦¦
¦¦+-----------+------------+-------+----------------------+------+------------+¦¦
¦¦¦ DESCRIÇÃO ¦ Relatorio de gasto de frete.                                  ¦¦¦
¦¦+-----------+---------------------------------------------------------------+¦¦
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
+-------------------------------------------------------------------------------+
/*/
User Function CTBGASTO()
    
    Local oTela
    Local dtATE := STOD("")
    Local dtDE  := STOD("")
		
	// WfPrepEnv('01', '01', "U_xRELCUSTO",, "SIGAFIN")
	@ 200,1 TO 380,380 DIALOG  oTela TITLE OemToAnsi("Sobre gasto e saldo")
	@ 02,010 TO 080,180
	@ 10,018 Say "Relatório Sobre gasto e saldo"
    @ 22,018 say "Data de:"
    @ 40,018 say "Data ate:"
	@ 20,055 MsGet dtDE Size 50,10 Of oTela Pixel
	@ 38,055 MsGet dtATE Size 50,10 Of oTela Pixel 
	@ 60,118 BMPBUTTON TYPE 01 ACTION Processa({||execRel(dtDE,dtATE),"Processando Relatorio"})
	@ 60,148 BMPBUTTON TYPE 02 ACTION Close( oTela)   
	Activate Dialog  oTela Centered
Return 
Static Function execRel(dtDE,dtATE)
    Local cQry :=""
    Local oFWMsExcel
 	Local cPath := "C:\Relatorio\"  
 	Local cNomArq := "Sobre_gasto_saldo.xls"
	Local cArquivo := cPath + cNomArq
	MAKEDIR(cPath)
	
	If file(cArquivo)
		fErase(cArquivo)
	EndIf
    cQry += "exec sp_TEMP_GASTO '"+  DTOS(dtDE) + "','" +DTOS(dtATE) +"'"
    
    TCQuery cQry New Alias "TMP"
    // CSetField("TMP", "DATADEF", "D")
    oFWMsExcel := FwMsExcelEx():New()     
    oFWMsExcel:AddworkSheet("SOBRE_GASTO_SALDO") //Não utilizar número junto com sinal de menos. Ex.: 1-
        //Criando a Tabela
        oFWMsExcel:AddTable("SOBRE_GASTO_SALDO","Sobre Gasto e Saldo" )
        //Criando Colunas
        oFWMsExcel:AddColumn("SOBRE_GASTO_SALDO","Sobre Gasto e Saldo","CODIGO"		     ,1,1) 
        oFWMsExcel:AddColumn("SOBRE_GASTO_SALDO","Sobre Gasto e Saldo","QTD_FRETE"	     ,1,1) 
        oFWMsExcel:AddColumn("SOBRE_GASTO_SALDO","Sobre Gasto e Saldo","SOBREFRETE"	     ,1,1) 
        oFWMsExcel:AddColumn("SOBRE_GASTO_SALDO","Sobre Gasto e Saldo","QTDSD3"	         ,1,1) 
        oFWMsExcel:AddColumn("SOBRE_GASTO_SALDO","Sobre Gasto e Saldo","QDTSALDO"	     ,1,1) 
        oFWMsExcel:AddColumn("SOBRE_GASTO_SALDO","Sobre Gasto e Saldo","CONSUMOACM"      ,1,1) 
        
        //Criando as Linhas
        While !(TMP->(EoF()))
            oFWMsExcel:AddRow("SOBRE_GASTO_SALDO","Sobre Gasto e Saldo",{;
                TMP->(CODIGO	   ),;
				TMP->(QTD_FRETE    ),;
				TMP->(SOBREFRETE   ),;
				TMP->(QTDSD3       ),;
				TMP->(QTDSALDO     ),;
				TMP->(CONSUMOACM   );
            })
        TMP->(DbSkip())
        EndDo
    TMP->(dbCloseArea())
    oFWMsExcel:Activate()
    oFWMsExcel:GetXMLFile(cArquivo)
    //Abrindo o excel e abrindo o arquivo xml
   oExcel := MsExcel():New()                //Abre uma nova conexão com Excel
   oExcel:WorkBooks:Open(cArquivo)          //Abre uma planilha
   FWAlertSuccess('<ATENÇÃO> Caso a Planilha não abra automaticamente,a mesma poderá ser encontrada em ' + cPath ,'Relatório gerado')
   oExcel:SetVisible(.T.)                  //Visualiza a planilha           
//    oExcel:Destroy()
Return 
