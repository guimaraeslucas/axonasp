# Instruções rápidas para agentes de código

Objetivo: ajudar um agente a ser produtivo rapidamente neste repositório (interpretador ASP em Go). Você é um desenvolvedor GoLang experiente, focado em qualidade, precisão, performance e segurança.

- **Arquitetura (visão geral)**
  - O servidor HTTP principal está em `main.go`. Ele escuta em `:4050` e serve a pasta `./www`.
  - A pasta `asp/` contém o interpretador:
    - `asp_parser.go`.: parser ASP que divide o código em blocos HTML e ASP, e faz parse dos blocos ASP via VBScript-Go.
    - `asp_lexer.go`: lexer ASP que tokeniza o código em blocos (HTML, ASP, directives). Suporta diretivas ASP como `<%@ Language=VBScript %>`.
 - A pasta `server/libs/` contém implementações de bibliotecas ASP (FileSystem, JSON, HTTP, Mail, Template, Crypto) que devem ser integradas ao servidor ASP e que são chamadas através de `Server.CreateObject("G3NOME_DA_BIBLIOTECA")`. Elas possuem um execution context que permite acesso ao contexto ASP atual (Request, Response, Session, Application, Server). 
  - Configurações do programa/bibliotecas devem ser obtidas de um arquivo .env e devem ter opções padrão no código caso .env não seja obtido
  - A pasta `server/deprecated/` contém implementações antigas que não devem ser modificadas. Utilize apenas para referência.

- **Como rodar / depurar localmente**
  - Com Go: `go run main.go` (porta padrão `4050`).
  - Para build: `go build -o go-asp.exe` e executar `./go-asp.exe` (Windows Powershell).
  - O nosso terminal roda Windows Powershell, então utilize comandos compatíveis.
  - Testes manuais: abra `http://localhost:4050/test_basics.asp` ou outros `www/test_*.asp`.
  - Para ver trace de panic detalhado (stacktrace HTML), defina a variável ASP `debug_asp_code` para "TRUE" no próprio ASP antes do erro, por exemplo:
    - `<% debug_asp_code = "TRUE" %>` no topo do arquivo ASP.

- **Padrões do projeto e convenções**
  - ASP directives like `<%@ Language=VBScript %>` are supported. They are parsed by the lexer and stored as directive blocks with attributes.
  - CÓDIGO E COMENTÁRIOS DEVEM **OBRIGATORIAMENTE** ESTAR EM INGLÊS
  - Lookup de variáveis é case-insensitive; internamente as chaves são armazenadas em minúsculas.
  - Includes: `<!--#include file="..."-->` é relativo ao arquivo atual; `virtual` é relativo a `www/` (root).
  - `Session`  deve ser armazenado em uma pasta em `temp/session` e e `Application` na memória; a sessão usa cookie `ASPSESSIONID`.
  - Comentários, nomes de funções, variáveis, campos e descritores devem obrigatoriamente ser em Inglês Americano
  - Sempre que possível utilize as funções já disponíveis em GoLang ou em bibliotecas GoLang, quando adequado, faça apenas a tradução do código ASP Clássico para GoLang. Nunca modifique a forma com que o usuário precisa entrar com os dados da função, use o modelo ASP Clássico, adapte o processamento para GoLang e dê o retorno como se fosse ASP. Esse é um interpretador em tempo real.
  - Siga o padrão e a conformidade estrita de VBScript.
  - Trabalhamos por padrão com UTF-8
  - Nas suas ações como Agente, ao explicar, seja sucinto e direto nas explicações, de forma a ser mais rápido e utilizar menos texto. Não fique criando arquivos .md de instruções extras, utilize este arquivo para todas as instruções úteis ao Agente.
  - Ao atualizar copilot-instructions.md, atualize também o arquivo GEMINI.md e vice-versa.
  - Comentários devem OBRIGATÓRIAMENTE ser escritos em inglês, se você econtrar comentários em outra língua ou em português, traduza-os IMEDIATAMENTE para inglês.
  - As páginas de modelo test_*.asp e a página Default.asp devem ser todas obrigatóriamente em inglês, caso você encontre outro idioma, traduza para inglês.
  - Legacy ASP functions and behaviors must be strictly followed to ensure compatibility with existing ASP codebases.

- **Boas mudanças para PRs pequenas**
  - Fixes e features pontuais devem incluir um exemplo ASP em `www/` (ou atualizar `test_*.asp`) e adicionar a modificação um link ou atualizar a descrição em `Default.asp`, sempre seguindo uma formatação adequada. Não crie arquivos .MD com a descrição das funcionalidades. Tudo deve estar em inglês americano..
  - Evitar reescrever estilo de arquivos; manter implementações pequenas e testáveis.
  - Não modificar nada que esteja na pasta VBScript-Go.
  - Evitar modificar o nome do programa que é G3 AxonASP
  - Quando criar bibliotecas novas, siga o padrão de nomenclatura `*_lib.go` (ex.: `string_lib.go`). Procure manter o máximo de similaridade com VBscript na nomenclatura das funções e forma de chamadas. Documente as funções novas com comentários no padrão GoLang e em inglês. As libs são chamadas através de Server.CreateObject("G3NOME_DA_BIBLIOTECA"), se necessário aprenda lendo uma lib de exemplo para entender a chamada e a interface.
  - Siga padrões de segurança avançados e atuais, por exemplo, dando a opção do usuário usar variáveis de ambiente para configuração por exemplo.


- **Arquivos-chave para referência rápida**
  - `main.go` — servidor e fluxo principal
  - `server/libs/*_lib.go` — implementações de bibliotecas auxiliares (File, HTTP, JSON, Mail, Template)
  - `www/` — páginas ASP de exemplo e testes

- **Configuração via .env**
  - O servidor suporta um arquivo `.env` na raiz para configuração.
  - Variáveis:
    - `SERVER_PORT`: Porta HTTP (padrão 4050)
    - `WEB_ROOT`: Diretório raiz dos arquivos ASP (padrão ./www)
    - `TIMEZONE`: Fuso horário (padrão America/Sao_Paulo)
    - `DEFAULT_PAGE`: Página inicial (padrão default.asp)
    - `SCRIPT_TIMEOUT`: Timeout em segundos (padrão 30)
    - `DEBUG_ASP`: Ativa modo Debug para parser ASP/VB (padrão FALSE). Quando TRUE, cada erro, notice e warning do parse são notificados no console com detalhes como número de linha.
    - `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASS`, `SMTP_FROM`: Configurações de envio de e-mail (Mail.SendStandard).

- **Compilação obrigatória**
  - Ao terminar a edição do código executável GoLang, sempre compile o programa em GoLang para a plataforma Windows e veja se foi compilado com sucesso. Não é necessário iniciar o programa a não ser que o usuário precise ver a modificação em ação.
  - Caso o usuário peça a edição de arquivos ASP, não é necessário compilar o programa em GoLang, apenas editar o arquivo ASP solicitado.

- **Funções VBScript Implementadas** 
  - **Libraries (via Server.CreateObject)**:
    - `G3JSON`: NewObject, Parse, Stringify, LoadFile
    - `G3FILES`: Read, Write, Append, Exists, Size, List, Delete, MkDir
    - `G3HTTP`: Fetch (method)
    - `G3TEMPLATE`: Render
    - `G3MAIL`: Send, SendStandard
    - `G3CRYPTO`: UUID, HashPassword, VerifyPassword
  - **COM Standard**: `MSXML2.XMLHTTP`, `MSXML2.ServerXMLHTTP`, `MSXML2.DOMDocument`, `ADODB.Connection`, `ADODB.Recordset`, `ADODB.Stream`
  - **MSXML2 Objects** (`msxml_lib.go`):
    - `MSXML2.ServerXMLHTTP`: HTTP requests with XML support (Open, SetRequestHeader, Send, GetResponseHeader, GetAllResponseHeaders, Status, StatusText, ResponseText, ReadyState, Timeout)
    - `MSXML2.DOMDocument`: XML DOM manipulation (LoadXML, Load, Save, GetElementsByTagName, SelectSingleNode, SelectNodes, CreateElement, CreateTextNode, CreateAttribute, AppendChild, ParseError)
  - **ADODB Database Support**:
    - `ADODB.Connection`: Open, Close, Execute, BeginTrans, CommitTrans, RollbackTrans, Errors (collection)
    - `ADODB.Recordset`: Open, Close, MoveNext, MoveFirst, MoveLast, MovePrevious, AddNew, Update, Delete, EOF, BOF, RecordCount, Fields, Supports (method), Filter (property)
    - `ADODB.Stream`: Open, Close, Read, ReadText, Write, WriteText, LoadFromFile, SaveToFile, CopyTo, Flush, SetEOS, SkipLine, Type, Mode, State, Position, Size, Charset, LineSeparator, EOS
    - **Connection.Errors**: Collection with Count, Item, Clear methods for error tracking
    - **Recordset.Supports**: Checks provider capabilities (adAddNew, adUpdate, adDelete, adMovePrevious, adFind, etc.)
    - **Recordset.Filter**: In-memory filtering with operators: =, <>, >, <, >=, <=, LIKE
    - Supported Databases: SQLite (in-memory and file), MySQL, PostgreSQL, MS SQL Server
    - Connection String Formats:

      ```
      
      sqlite::memory:
      sqlite:./mydata.db
      Driver={MySQL ODBC Driver};Server=localhost;Database=mydb;UID=root;PWD=password
      Driver={PostgreSQL ODBC Driver};Server=localhost;Database=mydb;UID=postgres;PWD=password;Port=5432
      Driver={ODBC Driver 17 for SQL Server};Server=localhost;Database=mydb;UID=sa;PWD=password;Port=1433
      
      ```

    - Field Access: rs("fieldname") or rs.Fields.Item("fieldname") or rs.Item("fieldname")


Se algo ficou obscuro ou você quer que eu detalhe um ponto (por exemplo, como adicionar suporte a um objeto COM ou persistir sessões), diga qual área quer priorizar.
