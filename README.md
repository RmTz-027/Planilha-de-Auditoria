# Auditoria de Estoque - Planilha Automatizada

Este projeto foi desenvolvido para auxiliar no processo de auditoria de estoque utilizando uma planilha do Google Sheets integrada com uma interface personalizada em HTML, CSS e Google Apps Script.

### Contexto

No início, as auditorias eram realizadas manualmente. Consultávamos o endereço (local onde o produto está armazenado) diretamente no ERP — neste caso, o **GK e-Manager** — e contávamos "a dedo" os produtos.

A primeira versão do projeto usava fórmulas em uma planilha: após baixarmos um relatório do GK e-Manager com todo o estoque, filtrávamos o endereço desejado e, ao bipar o código da mercadoria, a planilha exibia se o produto pertencia àquele endereço.

Com o tempo, várias fórmulas começaram a ser alteradas indevidamente, então iniciei o estudo de Google Apps Script e ,usando também HTML e CSS, desenvolvi uma interface que facilita a visualização das informações e evita erros dentro da planilha.

---

## Funcionalidades

- Obtém os dados existentes na aba "BANCO DE DADOS";
- Gera a lista de produtos pertencente ao endereço selecionado;
- Verifica se o código informado pertence ao endereço selecionado;
- Adiciona novos dados à aba "BANCO DE DADOS";
- Salva os códigos inseridos em uma aba;
- Envia dados para outra planilha em caso de divergências graves;

---

## Tecnologias utilizadas

- Google Apps Script
- HTML5
- CSS3
- Google Sheets
- Publicação como App da Web

---

## Estrutura dos arquivos


auditoria-estoque/ 
├── index.html # Interface da aplicação 
├── funcoes.gs # Código em Google Apps Script
└── README.md # Documentação do projeto

---

## Como usar

1. Faça uma cópia da planilha original no seu Google Drive.
2. Vá em **Extensões > Apps Script** e substitua o código existente pelo conteúdo do arquivo `funcoes.gs`.
3. No editor do Apps Script, adicione os arquivos HTML e CSS, colando o conteúdo de `index.html`.
4. Salve o código.
5. Vá em **Implementar > Nova Implementação > Seleciona o Tipo > App da Web > Preencha o Necessário**.
6. Utilize o link gerado para acessar a interface personalizada.

---

## 📷 Preview

![Interface da Auditoria]


## Autor

Desenvolvido por [Ramon Madeira](https://github.com/RmTz-027)  


