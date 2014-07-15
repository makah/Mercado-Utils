Attribute VB_Name = "UpdateValues"
'---------------------------------------------------------'
' Sistema para exportar os valores para o PerformIt
'
' @author Mauricio Arieira
' @company Oceana Investimentos
' @date 09/07/2014
'---------------------------------------------------------'

'----------------------- Constante -----------------------'
'PerformIt
Const PERFORMIT_URL As String = "https://mycompany.performit.com.br/quote/publishFromExcel"
Const PERFORMIT_KEY As String = "Hk0nEvagQTo5"

'Sistema
Const RUNNING As String = "Rodando"
Const STOPPED As String = "Parado"

Const COLUMN_ROWINDEX = 3
Const SCHEDULER_INTERVAL As Integer = 30
'---------------------------------------------------------'

'------------------------ Globais ------------------------'

'O tempo que o scheduler irá executar a função novamente.
'Essa informação precisa ser global para conseguirmos cancelar
'o scheduler.
Dim scheduleTime As Date


'A sheet que será passada para a função a ser executada.
'
'Essa informação é importante porque o scheduler ficará rodando
'e se algum outro Excel for aberto na maquina, a macro vai parar
'de funcionar corretamente, podendo até sobrescrever dados de
'outras planilhas abertas.
Dim dataSheet As Worksheet
'---------------------------------------------------------'


'Toggle Button
Sub Execute()
Attribute Execute.VB_ProcData.VB_Invoke_Func = " \n14"
    Set dataSheet = ActiveSheet

    If [UPDATE_STATUS] Like RUNNING Then
        Call StopScheduler
    Else
        Call RunScheduler
    End If
End Sub

'Inicia o Scheduler
'@param sheet
Private Sub RunScheduler()
    [UPDATE_STATUS] = RUNNING
    scheduleTime = Now + TimeSerial(0, 0, SCHEDULER_INTERVAL)
    
    Call UpdateValues(Array(1, 3, 4))
    
    Application.OnTime EarliestTime:=scheduleTime, procedure:="RunScheduler", Schedule:=True
End Sub

'Para o Scheduler
Private Sub StopScheduler()
    [UPDATE_STATUS] = STOPPED
    
    On Error GoTo ERRORHANDLER:
    Application.OnTime EarliestTime:=scheduleTime, procedure:="RunScheduler", Schedule:=False
    On Error GoTo 0
    
    Exit Sub
    
ERRORHANDLER:
    Debug.Print "[Error] StopScheduler: " & Err.Description
End Sub

'Envia os dados para o PerformIt
'@param cols Colunas que serão enviadas para o performIt
Private Sub UpdateValues(ByRef cols As Variant)
    Dim myRange, lastRow As Integer
    Dim col, colHeader As String
    Dim url As String, urlParam As String
    Dim symbol As String, price As String, change As String
    
    With dataSheet
        For Each col In cols
            lastRow = .Cells(100000, col).End(xlUp).Row
        
            If colHeader = vbNullString Then
                colHeader = "?" & .Cells(COLUMN_ROWINDEX, col) & "="
            Else
                colHeader = "&" & .Cells(COLUMN_ROWINDEX, col) & "="
            End If
            myRange = .Range(.Cells(COLUMN_ROWINDEX + 1, col), .Cells(lastRow, col))
            urlParam = urlParam & colHeader & Join(WorksheetFunction.Transpose(myRange), "_")
        Next col
    End With
    
    urlParam = urlParam & "&timestamp=1&pricingApiKey=" & PERFORMIT_KEY
    url = PERFORMIT_URL & Replace(urlParam, ",", ".")

    'Debug.Print "URL: " & url
    [HTTP_RESPONSE] = PostRequest(url)
    
    'Funcao que atualiza os valores do Excel.
    'Fiz o update depois de atualizar para dar tempo de todos os precos serem atualizados
    '  antes de enviarmos para o sistema
    Application.Run "RefreshAllWorkbooks"
End Sub

'Executa um post com a URL
'@param url A url que será enviada
Private Function PostRequest(url) As String
    Dim toResolve As Integer, toConnect As Integer, toSend As Integer, toReceive As Integer
    
    Set XMLHTTP = CreateObject("Msxml2.ServerXMLHTTP.6.0")
    
    'setTimeout
    toResolve = 10 * 1000
    toConnect = 10 * 1000
    toSend = 10 * 1000
    toReceive = 15 * 1000
    XMLHTTP.setTimeouts toResolve, toConnect, toSend, toReceive
    XMLHTTP.setOption 2, 13056
    
    'POST
    XMLHTTP.Open "POST", url, False
    XMLHTTP.send
    
    PostRequest = XMLHTTP.responseText
End Function
