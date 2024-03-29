#INCLUDE 'PROTHEUS.CH'
#INCLUDE 'TOTVS.CH'

User Function konInv01()
Local bSair 
Private oDlg  := Nil

    bSair := { || oDlg:DeActivate()}

    oDlg  := FwDialogModal():New()       
    oDlg:setBackground(.T.)
    oDlg:SetTitle('Importar Arquivo xls')
    oDlg:EnableFormBar(.T.)
    oDlg:SetEscClose(.T.)
    oDlg:SetCloseButton(.F.)

    //Seta a largura e altura da janela em pixel
    oDlg:SetPos(000, 000)
    oDlg:SetSize(200, 200)
    oDlg:CreateDialog()
    oDlg:CreateFormBar()

    @ 005,004 GROUP ogroup TO 160,180 PROMPT 'Observac�es ...'

    // oDlg:AddCloseButton(Nil,'Fechar')
    oDlg:AddButton('Abrir Arquivo' ,{|| InvFile()},,,.T.,.F.)
    oDlg:AddButton('Sair' ,bSair,,,.T.,.T.,)

    oDlg:GetPanelMain()
    oDlg:Activate()

Return

/*/{Protheus.doc} User Function konInv01 /*/

Static Function InvFile()

    Local aArea   := GetArea()
    Local cDirIni := GetTempPath()
    Local cTipArq := "Todas extens�es (*.*) | Arquivos texto (*.txt) | Arquivos com separa��es (*.csv)"
    Local cTitulo := "Sele��o de Arquivos para Processamento"
    Local lSalvar := .F.
    Local cArqSel := ""
    Local cArqCSV := ""
    
    //Se n�o estiver sendo executado via job
    If ! IsBlind()
 
        //Chama a fun��o para buscar arquivos
        cArqSel := tFileDialog(;
            cTipArq,;  // Filtragem de tipos de arquivos que ser�o selecionados
            cTitulo,;  // T�tulo da Janela para sele��o dos arquivos
            ,;         // Compatibilidade
            cDirIni,;  // Diret�rio inicial da busca de arquivos
            lSalvar,;  // Se for .T., ser� uma Save Dialog, sen�o ser� Open Dialog
            ;          // Se n�o passar par�metro, ir� pegar apenas 1 arquivo; Se for informado GETF_MULTISELECT ser� poss�vel pegar mais de 1 arquivo; Se for informado GETF_RETDIRECTORY ser� poss�vel selecionar o diret�rio
        )
 
        If ! Empty(cArqSel) .and. File(cArqSel)
            MsgInfo("O arquivo selecionado foi: " + cArqSel, "Aten��o")
            cArqCSV := fXLStoCSV(cArqSel)
            SLEEP(5000)
            if File(cArqCSV)
                Processa({|| fImporta(cArqCSV) }, 'Importando...')
                FWAlertSuccess("Arquivo importado com sucesso", "Success")
                fErase(cArqCSV) // DELETA A O ARQUIVO CSV
            EndIf
        EndIf
    EndIf
 
    RestArea(aArea)
    oDlg:DeActivate()

Return 

/*/ {Protheus.doc} fImporta(cArqCSV) /*/

Static Function fImporta(cArqSel)
    Local cDirTmp    := GetTempPath()
    Local cArqLog    := 'importacao_' + dToS(Date()) + '_' + StrTran(Time(), ':', '-') + '.log'
    Local nTotLinhas := 0
    Local cLinAtu    := ''
    Local nLinhaAtu  := 0
    Local aLinha     := {}
    Local oArquivo
    Local cPastaErro := '\x_logs\'
    Local cNomeErro  := ''
    Local cTextoErro := ''
    Local aLogErro   := {}
    Local nLinhaErro := 0
    Local cLog       := ''
    //Vari�veis do ExecAuto
    Private aDados         := {}
    Private lMSHelpAuto    := .T.
    Private lAutoErrNoFile := .T.
    Private lMsErroAuto    := .F.
    //Vari�veis da Importa��o
    Private cAliasImp  := 'SBM'
    Private cSeparador := ';'
 
    //Abre as tabelas que ser�o usadas
    DbSelectArea(cAliasImp)
    (cAliasImp)->(DbSetOrder(1))
    (cAliasImp)->(DbGoTop())
 
    //Definindo o arquivo a ser lido
    oArquivo := FWFileReader():New(cArqSel)
 
    //Se o arquivo pode ser aberto
    If (oArquivo:Open())
 
        //Se n�o for fim do arquivo
        If ! (oArquivo:EoF())
 
            //Definindo o tamanho da r�gua
            aLinhas := oArquivo:GetAllLines()
            nTotLinhas := Len(aLinhas)
            ProcRegua(nTotLinhas)
 
            //M�todo GoTop n�o funciona (dependendo da vers�o da LIB), deve fechar e abrir novamente o arquivo
            oArquivo:Close()
            oArquivo := FWFileReader():New(cArqSel)
            oArquivo:Open()
 
            //Iniciando controle de transa��o
            Begin Transaction
 
                //Enquanto tiver linhas
                While (oArquivo:HasLine())
 
                    //Incrementa na tela a mensagem
                    nLinhaAtu++
                    IncProc('Analisando linha ' + cValToChar(nLinhaAtu) + ' de ' + cValToChar(nTotLinhas) + '...')
 
                    //Pegando a linha atual e transformando em array
                    cLinAtu := oArquivo:GetLine()

                    If !nLinhaAtu == 1
                        aLinha := Separa(cLinAtu, cSeparador)
                    EndIf
 
                    //Se houver posi��es no array
                    If  !empty(aLinha) .AND. Len(aLinha) > 0
                        aDados := {}
                        aAdd(aDados, {'B7_FILIAL', CFilant, Nil})
                        aAdd(aDados, {'B7_COD', aLinha[2], Nil})
                        aAdd(aDados, {'B7_DESC', aLinha[3], Nil})
                        aAdd(aDados, {'B7_TIPO', aLinha[4], Nil})
                        aAdd(aDados, {'B7_LOCAL', aLinha[7], Nil})
                        aAdd(aDados, {'B7_QUANT', aLinha[8], Nil})
                        // aAdd(aDados, {'B7_DATA', aLinha[4], Nil})
 
                        lMsErroAuto := .F.
                        // MSExecAuto({|x, y| MATA270(x, y)}, aDados, 3)
 
                        //Se houve erro, gera o log
                        If lMsErroAuto
                            cPastaErro := '\x_logs\'
                            cNomeErro  := 'erro_' + cAliasImp + '_' + dToS(Date()) + '_' + StrTran(Time(), ':', '-') + '.txt'
 
                            //Se a pasta de erro n�o existir, cria ela
                            If ! ExistDir(cPastaErro)
                                MakeDir(cPastaErro)
                            EndIf
 
                            //Pegando log do ExecAuto, percorrendo e incrementando o texto
                            aLogErro := GetAutoGRLog()
                            For nLinhaErro := 1 To Len(aLogErro)
                                cTextoErro += aLogErro[nLinhaErro] + CRLF
                            Next
 
                            //Criando o arquivo txt e incrementa o log
                            MemoWrite(cPastaErro + cNomeErro, cTextoErro)
                            cLog += '- Falha ao incluir registro, linha [' + cValToChar(nLinhaAtu) + '], arquivo de log em ' + cPastaErro + cNomeErro + CRLF
                        Else
                            cLog += '+ Sucesso no Execauto na linha ' + cValToChar(nLinhaAtu) + ';' + CRLF

                            // fErase(cArqCSV)
                        EndIf
 
                    EndIf
 
                EndDo
            End Transaction
 
            //Se tiver log, mostra ele
            If ! Empty(cLog)
                MemoWrite(cDirTmp + cArqLog, cLog)
                ShellExecute('OPEN', cArqLog, '', cDirTmp, 1)
            EndIf

        Else
            MsgStop('Arquivo n�o tem conte�do!', 'Aten��o')
        EndIf
 
        //Fecha o arquivo
    Else
        MsgStop('Arquivo n�o pode ser aberto!', 'Aten��o')
    EndIf

    oArquivo:Close()
    oArquivo:Erase()
    //oDlg:DeActivate()

Return 

//Essa fun��o foi baseada como refer�ncia no seguinte link: https://stackoverflow.com/questions/1858195/convert-xls-to-csv-on-command-line
Static Function fXLStoCSV(cArqXLS)
    Local cArqCSV    := ""
    Local cDirTemp   := GetTempPath()
    Local cArqScript := cDirTemp + "XlsToCsv.vbs"
    Local cScript    := ""
    Local cDrive     := ""
    Local cDiretorio := ""
    Local cNome      := ""
    Local cExtensao  := ""
 
    //Monta o Script para converter
    cScript := 'if WScript.Arguments.Count < 2 Then' + CRLF
    cScript += '    WScript.Echo "Error! Please specify the source path and the destination. Usage: XlsToCsv SourcePath.xls Destination.csv"' + CRLF
    cScript += '    Wscript.Quit' + CRLF
    cScript += 'End If' + CRLF
    cScript += 'Dim oExcel' + CRLF
    cScript += 'Set oExcel = CreateObject("Excel.Application")' + CRLF
    cScript += 'Dim oBook' + CRLF
    cScript += 'Set oBook = oExcel.Workbooks.Open(Wscript.Arguments.Item(0))' + CRLF
    cScript += 'oBook.SaveAs WScript.Arguments.Item(1), 6' + CRLF
    cScript += 'oBook.Close False' + CRLF
    cScript += 'oExcel.Quit' + CRLF
    MemoWrite(cArqScript, cScript)
 
    //Pega os detalhes do arquivo original em XLS
    SplitPath(cArqXLS, @cDrive, @cDiretorio, @cNome, @cExtensao)
 
    //Monta o nome do CSV, conforme os detalhes do XLS
    // cArqCSV := cDrive + cDiretorio + cNome + ".csv"
    cArqCSV := cDrive + cDiretorio + cNome + ".csv"
    
 
    //Executa a convers�o, exemplo: 
    //   c:\totvs\Testes\XlsToCsv.vbs "C:\Users\danat\Downloads\tste2.xls" "C:\Users\danat\Downloads\tst2_csv.csv"
    ShellExecute("OPEN", cArqScript, ' "' + cArqXLS + '" "' + cArqCSV + '"', cDirTemp, 0 )
 
Return cArqCSV
