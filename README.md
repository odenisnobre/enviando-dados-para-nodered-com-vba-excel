
<h1 align="center">
<br>
Enviando dados para Node-Red usando VBA com Excel
</h1>

<p align="center">O objetivo é demonstrar de forma bem simples como enviar dados para o node-red usando o vba do Excel.</p>

<p align="center">
  <a href="https://www.apache.org/licenses/LICENSE-2.0">
    <img src="https://img.shields.io/badge/apache-2.0-blue" alt="License MIT">
  </a>
</p>

<div>
  <img src="https://github.com/dedynobre/enviando-dados-para-nodered-com-vba-excel/blob/master/func.gif" alt="vba-node-red" height="425">
</div>

<hr />

## Recursos utilizados

- **Node-Red** - Versão 1.1.0
- **Excel** - Versã 2016 x64

## Desenvolvimento

#### Configuração Node-Red
A configuração do node-red é bem simples, usamos um um node http request, um node http response um debug para poder visualizar os dados recebido.
O método utilizado para a requisição http foi o *POST*.
<div>
  <img src="https://github.com/dedynobre/enviando-dados-para-nodered-com-vba-excel/blob/master/ndr1.png" alt="vba-node-red" height="200">
</div>

#### Configuração VBA
```
Attribute VB_Name = "Módulo1"
Sub teste()

Dim hReq As Object
Dim i As Long

'url do caminho http que irá receber a requisição
strUrl = "http://localhost:1880/excel"

'célula que contém os dados a serem enviados
dados = Range("C3").Value

'configuração a conexão
Set hReq = CreateObject("MSXML2.XMLHTTP")
    With hReq
        .Open "POST", strUrl, False
        .Send (dados) 'envia os dados
    End With


response = hReq.ResponseText

End Sub

```

Caso a orgem da requisição fosse ao contrátio, ou seja, ao invés de enviar dados(POST) gostaria de receber dados seria necessário fazer algumas mudanças.
- Alterar o node http request para o método GET:
<div>
  <img src="https://github.com/dedynobre/enviando-dados-para-nodered-com-vba-excel/blob/master/ndr2.png" alt="vba-node-red" height="200">
</div>

- Alterar código vba para o cógido abaixo:
```
Attribute VB_Name = "Módulo1"
Sub teste()

Dim hReq As Object
Dim i As Long

'url do caminho http que irá receber a requisição
strUrl = "http://localhost:1880/excel"

'célula que contém os dados a serem enviados
dados = Range("C3").Value

'configuração a conexão
Set hReq = CreateObject("MSXML2.XMLHTTP")
    With hReq
        .Open "POST", strUrl, False
        .Send
    End With


response = hReq.ResponseText

range("C3").value = response

End Sub

```






## License

Para ver detalhes da licença, clique [Aqui](https://www.apache.org/licenses/LICENSE-2.0).

## Help

Caso precisem te ajuda ou tenham alguma sugestão, deixe seu comentário [Aqui](https://github.com/dedynobre/enviando-dados-para-nodered-com-vba-excel/issues).
