#Include "Protheus.ch" 
#include "rwmake.ch"                                            
#Include "TbiConn.ch"
#Include 'TOTVS.ch'
/*_______________________________________________________________________________
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¦¦+-----------+------------+-------+----------------------+------+------------+¦¦
¦¦¦ Programa  ¦ XRELCUSTO  ¦ Autor ¦ Edson Sales          ¦ Data ¦ 06/10/2022 ¦¦¦
¦¦+-----------+------------+-------+----------------------+------+------------+¦¦
¦¦¦ Descriçäo ¦ Relatorio Custo em Partes.                                    ¦¦¦ 
¦¦+-----------+---------------------------------------------------------------+¦¦
¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯*/

User Function XRELCUSTO()

	Local oTela,oEdit1
	Local nMes,nAno

	@ 200,1 TO 380,380 DIALOG  oTela TITLE OemToAnsi("Custo em Partes")
	@ 02,010 TO 080,180
	@ 10,018 Say "Relatório custo em Partes"
    @ 22,018 say "Mes:"
    @ 40,018 say "Ano:"
	@ 20,035 COMBOBOX oEdit1 VAR nMes ITEMS fMesRef() SIZE 30, 013 OF oTela PIXEL COLORS 0, 16777215
	@ 38,035 COMBOBOX oEdit1 VAR nAno ITEMS fAnoRef() SIZE 40, 013 OF oTela PIXEL COLORS 0, 16777215

	@ 60,118 BMPBUTTON TYPE 01 ACTION Processa({||consCusto(nMes,nAno),"Processando Relatorio"})
	@ 60,148 BMPBUTTON TYPE 02 ACTION Close( oTela)   
	//@ 70,100 BMPBUTTON TYPE 05 ACTION Pergunte(cPerg,.T.)

	Activate Dialog  oTela Centered

Return

Static Function consCusto(nMes,nAno)
	Local cQry  :=""
	Local cQry1 :=""
	Local cQry2 :=""
	Local cQry3 :=""
	Local cQry4 :=""
	Local cQry5 :=""
	Local nPeriodo
	Local oFWMsExcel
 	Local cPath := "C:\Custo_partes\"//   "\system\senf.xls"
 	Local cNomArq := "custo_em_partes.xls"
	Local cArquivo := cPath + cNomArq

	MAKEDIR(cPath)
	
	If file(cArquivo)
		fErase(cArquivo)
	EndIf
	nPeriodo := nAno+nMes
	cQry1 += " EXEC sp_TEMPSD3MES '"+nPeriodo+"'"+",'"+cfilant+"';"
	TCSqlExec(cQry1)

//************************ MONTA A TEMP PRODUCAO MES *******************************
//***************GERA O RELATORIO SINTETICO DE PRODUCAO MENSAL**********************
 	cQry2 += " EXEC sp_TEMPSD3MES_01; "
	TCSqlExec(cQry2)
	
 	cQry3 += " EXEC sp_TEMPSD3MES_02 '"+nPeriodo+"'"+",'"+cfilant+"';"    // + Substr(DTOS(dData),0,6)+"'"+",'"+cfilant+"';"
	TCSqlExec(cQry3)

 	cQry4 += " EXEC sp_TEMPSD3MES_03; "
	TCSqlExec(cQry4)

 	cQry5 += " EXEC sp_TEMPSD3MES_04 ; "
	TCSqlExec(cQry5)
 
                           										//  --RELATORIO APOS EXECUTAR TODAS AS PROCEDURES

 	cQry += " SELECT * FROM ##TEMPSD3MES_04 "
	cQry := ChangeQuery(cQry)

	dbUseArea(.T.,"TOPCONN",tcGenQry(,,cQry),"cAliasCust",.F.,.F.)
	DbSelectArea('cAliasCust')

																// oReport := Relcust(cAliasCust)
																// oReport:PrintDialog()

	oFWMsExcel := FwMsExcelEx():New()     
    oFWMsExcel:AddworkSheet("Custo") //Não utilizar número junto com sinal de menos. Ex.: 1-
        //Criando a Tabela
        oFWMsExcel:AddTable("Custo","Custos em Partes")
        //Criando Colunas
        oFWMsExcel:AddColumn("Custo","Custos em Partes","ANOME"			 ,1,1) 
        oFWMsExcel:AddColumn("Custo","Custos em Partes","COD"			 ,1,1) 
        oFWMsExcel:AddColumn("Custo","Custos em Partes","PRODUCAO"		 ,1,1) 
        oFWMsExcel:AddColumn("Custo","Custos em Partes","CUSTO UNIT"	 ,1,1) 
        oFWMsExcel:AddColumn("Custo","Custos em Partes","MP PA TOTAL"	 ,1,1) 
        oFWMsExcel:AddColumn("Custo","Custos em Partes","SCRAP"  		 ,1,1) 
        oFWMsExcel:AddColumn("Custo","Custos em Partes","MP PI TOTAL"	 ,1,1) 
        oFWMsExcel:AddColumn("Custo","Custos em Partes","T MO PI TOTAL"	 ,1,1) 
        oFWMsExcel:AddColumn("Custo","Custos em Partes","MO TOTAL"		 ,1,1) 
        oFWMsExcel:AddColumn("Custo","Custos em Partes","CUSTO TOTAL"	 ,1,1) 
        oFWMsExcel:AddColumn("Custo","Custos em Partes","MOD_PA"		 ,1,1) 
        oFWMsExcel:AddColumn("Custo","Custos em Partes","MOI_PA"	     ,1,1) 
        oFWMsExcel:AddColumn("Custo","Custos em Partes","GIF_PA"		 ,1,1) 
        oFWMsExcel:AddColumn("Custo","Custos em Partes","DESCRICAO"		 ,1,1) 
        oFWMsExcel:AddColumn("Custo","Custos em Partes","U MP PA"		 ,1,1) 
        oFWMsExcel:AddColumn("Custo","Custos em Partes","U SCRAP"		 ,1,1) 
        oFWMsExcel:AddColumn("Custo","Custos em Partes","U MP PI"		 ,1,1) 
        oFWMsExcel:AddColumn("Custo","Custos em Partes","U MOD PI"		 ,1,1) 
        oFWMsExcel:AddColumn("Custo","Custos em Partes","U MOD PA"		 ,1,1) 
        oFWMsExcel:AddColumn("Custo","Custos em Partes","U CUSTO UNIT"	 ,1,1) 
        oFWMsExcel:AddColumn("Custo","Custos em Partes","U MODS PA"		 ,1,1) 
        oFWMsExcel:AddColumn("Custo","Custos em Partes","U MOD PA1"	     ,1,1) 
        oFWMsExcel:AddColumn("Custo","Custos em Partes","U MOI PA1"	     ,1,1) 
        oFWMsExcel:AddColumn("Custo","Custos em Partes","U GIF PA1"	     ,1,1)

        //Criando as Linhas
        While !(cAliasCust->(EoF()))
            oFWMsExcel:AddRow("Custo","Custos em Partes",{;
                cAliasCust->(ANOMES	      ),;
				cAliasCust->(COD          ),;
				cAliasCust->(PRODUCAO	  ),;
				cAliasCust->(CUSTUNIT     ),;
				cAliasCust->(MP_PA_TOTAL  ),;
				cAliasCust->(SCRAP        ),;
				cAliasCust->(MP_PI_TOTAL  ),;
				cAliasCust->(T_MO_PI_TOTAL),;
				cAliasCust->(MO_TOTAL  	  ),;
				cAliasCust->(CUSTOTOTAL   ),;
				cAliasCust->(MOD_PA		  ),;
				cAliasCust->(MOI_PA		  ),;
				cAliasCust->(GIF_PA		  ),;
				cAliasCust->(DESCRICAO	  ),;
				cAliasCust->(U_MP_PA  	  ),;
				cAliasCust->(U_SCRAP  	  ),;
				cAliasCust->(U_MP_PI  	  ),;
				cAliasCust->(U_MOD_PI 	  ),;
				cAliasCust->(U_MOD_PA	  ),;
				cAliasCust->(U_CUSTUNIT   ),;
				cAliasCust->(MODS_PA      ),;
				cAliasCust->(U_MOD_PA1    ),;
				cAliasCust->(U_MOI_PA1    ),;
				cAliasCust->(U_GIF_PA1    );
            })
         cAliasCust->(DbSkip())
        EndDo 

		

    oFWMsExcel:Activate()
    oFWMsExcel:GetXMLFile(cArquivo)


    //Abrindo o excel e abrindo o arquivo xml
   oExcel := MsExcel():New()                //Abre uma nova conexão com Excel
   oExcel:WorkBooks:Open(cArquivo)          //Abre uma planilha
   FWAlertSuccess("Caso a Planilha não abra automaticamente,a mesma poderá ser encontrada em " + cPath ,"ATENÇÃO")
   oExcel:SetVisible(.T.)                   //Visualiza a planilha           
   oExcel:Destroy()
   cAliasCust->(DbCloseArea())

Return



Static Function fMesRef()
	Local aMesRef := {}
	
	aAdd(aMesRef , "01")
	aAdd(aMesRef , "02")
	aAdd(aMesRef , "03")
	aAdd(aMesRef , "04")
	aAdd(aMesRef , "05")
	aAdd(aMesRef , "06")
	aAdd(aMesRef , "07")
	aAdd(aMesRef , "08")
	aAdd(aMesRef , "09")
	aAdd(aMesRef , "10")
	aAdd(aMesRef , "11")
	aAdd(aMesRef , "12")

Return aMesRef

Static Function fAnoRef()
	Local aAnoRef := {}
	
	aAdd(aAnoRef , "2022")
	aAdd(aAnoRef , "2023")
	aAdd(aAnoRef , "2024")
	aAdd(aAnoRef , "2025")
	aAdd(aAnoRef , "2026")
	aAdd(aAnoRef , "2027")
	aAdd(aAnoRef , "2028")
	aAdd(aAnoRef , "2029")
	aAdd(aAnoRef , "2030")
	aAdd(aAnoRef , "2031")
	aAdd(aAnoRef , "2032")
	aAdd(aAnoRef , "2033")	
	aAdd(aAnoRef , "2034")
	aAdd(aAnoRef , "2035")
	aAdd(aAnoRef , "2036")
	aAdd(aAnoRef , "2037")
	aAdd(aAnoRef , "2038")
	aAdd(aAnoRef , "2039")
	aAdd(aAnoRef , "2040")
	aAdd(aAnoRef , "2041")
	aAdd(aAnoRef , "2042")
	aAdd(aAnoRef , "2043")
	aAdd(aAnoRef , "2044")
	aAdd(aAnoRef , "2045")
	aAdd(aAnoRef , "2046")
	aAdd(aAnoRef , "2047")
	aAdd(aAnoRef , "2048")
	aAdd(aAnoRef , "2049")

Return aAnoRef
