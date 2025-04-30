# vba_sap_qp02_apagar
Apagar análise, laboratório QP02

Sub ApagarCMQP02_Click()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Planilha1")

    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim totalLinhas As Integer
    Dim grid As Object
    
    Dim rng As Range
    Dim rngDeletar As Range
    Dim celula As Range
    Dim valorProcurado As String
    Dim encontrado As Boolean
    
    If Not IsObject(Appl) Then
        Set SapGuiAuto = GetObject("SAPGUI")
        Set Appl = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(Connection) Then
       Set Connection = Appl.Children(0)
    End If
    If Not IsObject(session) Then
       Set session = Connection.Children(0)
    End If
        
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n qp02"
    session.findById("wnd[0]").sendVKey 0
    
    For i = 6 To ultimaLinha
    
        On Error Resume Next
        
        session.findById("wnd[0]/usr/ctxtRC27M-MATNR").Text = ws.Cells(i, 1).Value '"51-01389"
        session.findById("wnd[0]/usr/ctxtRC27M-WERKS").Text = "BR35"
        session.findById("wnd[0]/usr/ctxtRC27M-WERKS").SetFocus
        session.findById("wnd[0]/usr/ctxtRC27M-WERKS").caretPosition = 4
        session.findById("wnd[0]").sendVKey 0
              
        '---------------------------------------------------------------------------------------------------------
        '---------------------------------------------------------------------------------------------------------
        '---------------------------------------------------------------------------------------------------------
        
        ' Espera os dados carregarem e lê os valores da primeira coluna da ALV (Grade)
        linha = 2
        indice = 0
        'erro = "________________________________________"
        

        Do While True
           
            'On Error GoTo FimDoLoop
            valorColuna = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6," & indice & "]").Text
            
            If valorColuna <> "" Then
                ws.Cells(linha, "D").Value = valorColuna
                linha = linha + 1
                indice = indice + 1
 
            Else
                Exit Do
                
            End If
        Loop
        
        
        
        
        Set rng = ws.Range("O2:O5") ' Altere conforme necessário - coluna fórmula "O"
         
        valorProcurado = "LABORATÓRIOFISICOQUÍMICO" ' Valor que você quer encontrar LABORATÓRIO FISICO-QUÍMICO

        encontrado = False
    
        For Each celula In rng
            
            'ws.Cells(i, "C").Value = celula.Value
            
            If StrComp(Trim(CStr(celula.Value)), Trim(valorProcurado), vbTextCompare) = 0 Then
                encontrado = True
                ws.Cells(i, "B").Value = "ACHOU"
                
                session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400").getAbsoluteRow(1).Selected = True ' melhorar para saber qual o index do item a ser excluido
                session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VORNR[0,1]").SetFocus
                session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VORNR[0,1]").caretPosition = 0
                session.findById("wnd[0]/tbar[1]/btn[14]").press
                session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                session.findById("wnd[0]/tbar[0]/btn[11]").press
                session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
                Application.Wait Now + TimeValue("00:00:02")
                Exit For
            End If
        Next celula
        
        
           
        If Not encontrado Then
            ws.Cells(i, "B").Value = "NÃO ACHOU"
            
            session.findById("wnd[0]/tbar[0]/okcd").Text = "/n qp02"
            session.findById("wnd[0]").sendVKey 0
            
            Application.Wait Now + TimeValue("00:00:02")
            
        End If
               
        '---------------------------------------------------------------------------------------------------------
        
        Set rngDeletar = ws.Range("D2:D5") ' coluna Valores "D"
        rng2.ClearContents ' Apaga os valores do range
        
        Application.Wait Now + TimeValue("00:00:02")
        
        statusMessage = session.findById("wnd[0]/sbar").Text
        ws.Cells(i, "G").Value = statusMessage
        
        typeMessage = session.findById("wnd[0]/sbar").messagetype
        ws.Cells(i, "H").Value = typeMessage
        
        If typeMessage <> "S" Then
            
            ws.Cells(i, "F").Value = "Não tem fisico-quimico"
            
            session.findById("wnd[0]/tbar[0]/okcd").Text = "/n qp02"
            session.findById("wnd[0]").sendVKey 0
            
        Else
            ws.Cells(i, "F").Value = "Sucesso"
            
        End If
        
    Next i
             
End Sub


