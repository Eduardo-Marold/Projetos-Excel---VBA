Attribute VB_Name = "CadastroMd"
Option Explicit

'Verifica se já existe um numero de celular cadastrado
Sub VerificaCelular(nCel As String, verCel As Integer)

    Dim wPlan   As Worksheet
    
    Set wPlan = Planilha4
    
    wPlan.Activate
    
    With Application.WorksheetFunction
        verCel = .CountIf(wPlan.Range("C:C"), nCel)
    End With
     
    Set wPlan = Nothing

End Sub


'Busca a próxima linha vazia para efetuar o cadastro
Sub UltimaLinha(ulin As Long, uval As Long)

    Dim wPlan   As Worksheet
    
    Set wPlan = Planilha4
    
    wPlan.Activate
    
    With wPlan.Application.WorksheetFunction
        uval = .Large(wPlan.Range("A:A"), 1)
        uval = uval + 1
        
        ulin = .Count(wPlan.Range("A:A"), 1)
        ulin = ulin + 1
    End With

    Set wPlan = Nothing

    'Do
    '    If Not (IsEmpty(ActiveCell)) Then
    '    ActiveCell.Offset(1, 0).Select
    '    End If
    'Loop Until IsEmpty(ActiveCell) = True
    
    'uLin = Cells.Find(What:="*", After:=Range("A1"), _
    '                SearchOrder:=xlByRows, _
    '                SearchDirection:=xlPrevious).Row
    'uLin = uLin + 1

End Sub


'Será usado para manipular as imagens de fundo
Sub SelecionaFundo(nmImg As String, img As String)

Dim wkb     As Workbook

Set wkb = Workbooks("CadastroClientes.xlsm")

img = wkb.Path & "\Fundos\" & nmImg & ".jpg"

End Sub


'Usado para criar cópias das imagens dentro de uma pasta especifica
'no primeiro com imagens, irá criar uma pasta de destino, caso a pasta já exista,
'irá apenas copiar as imagens para dentro da pasta

Sub CopiaPicClientes(Nmpic As String, nvDest As String, dstAnt As String)

''''''''Set wbk = Workbooks("CadastroClientes.xlsm")
''''''''
''''''''Nmpic = Format(txtCodigo.Value, "0") & ".jpg"
''''''''nvDest = wbk.Path & "\PicClientes"
''''''''dstAnt = txtImagem.Text
''''''''
''''''''Call CadastroMd.CopiaPicClientes(Nmpic, nvDest, dstAnt)

'Sub CopiaPicClientes()
'
'Dim Nmpic   As String
'Dim nvDest  As String
'Dim dstAnt  As String
'
'Nmpic = "imagem teste2.jpg"
'dstAnt = "C:\Users\Marcelo Nascimento\Pictures\invitation.jpg"
'nvDest = ThisWorkbook.Path & "\PicClientes"

On Error Resume Next

Dim arqImg As Scripting.FileSystemObject

Set arqImg = New Scripting.FileSystemObject

If Not arqImg.FolderExists(nvDest) Then
    arqImg.CreateFolder nvDest
End If

If arqImg.FileExists(dstAnt) Then
    arqImg.CopyFile _
        Source:=dstAnt, Destination:=nvDest & "\" & Nmpic
End If

Set arqImg = Nothing

End Sub

'Exclui imagens com base no endereço da imagem
Sub ExcluiPicClientes(DestArq As String)

On Error Resume Next

Dim arqImg As Scripting.FileSystemObject

Set arqImg = New Scripting.FileSystemObject

If Not arqImg.FolderExists(DestArq) Then
    arqImg.DeleteFile DestArq
End If

Set arqImg = Nothing

End Sub


Sub TrazDadosCnpj(Cnpj As String, nEmpres As String, nFantas As String, _
                    pAtivid As String, sAtiv1 As String, sAtiv2 As String, sAtiv3 As String, tpEmpre As String, _
                        logrado As String, nEndere As String, Complem As String, cepEnde As String, bairro As String, _
                        munic As String, ufEnd As String, Tel As String, email As String, situac As String, DtAbert As String)

'Dim Cnpj        As String
'Dim nEmpres     As String
'Dim nFantas     As String
'Dim pAtivid     As String
'Dim sAtiv1      As String
'Dim sAtiv2      As String
'Dim sAtiv3      As String
'Dim tpEmpre     As String
'Dim logrado     As String
'Dim nEndere     As String
'Dim Complem     As String
'Dim cepEnde     As String
'Dim bairro      As String
'Dim munic       As String
'Dim ufEnd       As String
'Dim Tel         As String
'Dim email       As String
'Dim situac      As String
'Dim DtAbert     As String
                        
Dim sHttp       As Object
Dim CvtJson     As Object

Dim vEntr       As String
Dim vSaid       As String
Dim vBusca      As String

On Error Resume Next

vEntr = Cnpj
vSaid = Replace(vEntr, "-", "")
vSaid = Replace(vSaid, ".", "")
vSaid = Replace(vSaid, "/", "")
vSaid = vSaid

'    Set sHttp = CreateObject("MSXML2.XMLHTTP")
'
'    With sHttp
'        .Open "GET", "http://receitaws.com.br/v1/cnpj/" & vBusca, False
'        .Send
'    End With

With Application.WorksheetFunction
    vBusca = .WebService("http://receitaws.com.br/v1/cnpj/" & vSaid)
End With

Set CvtJson = JsonConverter.ParseJson(vBusca)

'Cnpj = CvtJson("nome")
nEmpres = CvtJson("nome")
nFantas = CvtJson("fantasia")
'pAtivid = CvtJson("atividade_principal")(1)("text")
'sAtiv1 = CvtJson("atividades_secundarias")(1)("text")
'sAtiv2 = CvtJson("atividades_secundarias")(2)("text")
'sAtiv3 = CvtJson("atividades_secundarias")(3)("text")
tpEmpre = CvtJson("natureza_juridica")
logrado = CvtJson("logradouro")
nEndere = CvtJson("numero")
Complem = CvtJson("complemento")
cepEnde = CvtJson("cep")
bairro = CvtJson("bairro")
munic = CvtJson("municipio")
ufEnd = CvtJson("uf")
Tel = CvtJson("telefone")
email = CvtJson("email")
situac = CvtJson("situacao")
DtAbert = CvtJson("abertura")

End Sub

