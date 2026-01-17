# DEBUG ASP Parser - Guia de Uso

## Descrição

A funcionalidade de Debug permite que você ative mode de diagnóstico para o parser ASP/VB. Quando ativado, cada erro, notice e warning do parse serão notificados no console com detalhes como número de linha e mensagem de erro.

## Como Ativar

### Método 1: Variável de Ambiente

Defina a variável de ambiente `DEBUG_ASP` com o valor `TRUE`:

#### Windows PowerShell:
```powershell
$env:DEBUG_ASP = "TRUE"
.\asp-interpreter.exe
```

#### Windows CMD:
```cmd
set DEBUG_ASP=TRUE
asp-interpreter.exe
```

#### Linux/Mac:
```bash
export DEBUG_ASP=TRUE
./asp-interpreter
```

### Método 2: Arquivo .env

Crie um arquivo `.env` na raiz do projeto com o seguinte conteúdo:

```env
DEBUG_ASP=TRUE
```

Quando o servidor iniciar, ele carregará automaticamente esta configuração.

## Saída de Debug

Quando `DEBUG_ASP=TRUE` está ativado, a saída do console será similar a:

```
[DEBUG] DEBUG_ASP mode is enabled
[ASP Parser Error] Line 15: Variable 'myVar' is not defined
[ASP Parser Error] Line 23: Syntax error in expression
```

Cada erro listará:
- `[ASP Parser Error]` ou `[ASP Parse]` para informações gerais
- Número da linha onde o erro foi detectado
- Descrição detalhada do erro

## Exemplo Prático

### Arquivo ASP com Erro: `test.asp`
```asp
<%
  Dim x
  x = 10
  Response.Write(y) ' Erro: y não foi definido
%>
```

### Execução com Debug Desativado:
```
PS> .\asp-interpreter.exe
Starting G3pix AxonASP on http://localhost:4050
Serving files from ./www
```

### Execução com Debug Ativado:
```
PS> $env:DEBUG_ASP = "TRUE"
PS> .\asp-interpreter.exe
Starting G3pix AxonASP on http://localhost:4050
Serving files from ./www
[DEBUG] DEBUG_ASP mode is enabled
[ASP Parser Error] Line 4: Undefined variable 'y'
```

## Benefícios

- **Diagnóstico Rápido**: Identifique problemas de parsing rapidamente
- **Desenvolvimento**: Facilita o debug durante desenvolvimento de novas features
- **Troubleshooting**: Ajuda a diagnosticar problemas em arquivos ASP complexos
- **Produção**: Pode ser desativado facilmente para não poluir logs

## Dicas

1. Use `DEBUG_ASP=TRUE` durante desenvolvimento
2. Desative em produção para melhor performance
3. Combine com `SCRIPT_TIMEOUT` para debug de scripts longos
4. Verifique a linha indicada do erro - pode haver problemas anteriores que causam o erro reportado

## Configuração Padrão

Por padrão, `DEBUG_ASP` é definido como `FALSE`. Se a variável de ambiente não for definida e não houver arquivo `.env`, o modo debug permanecerá desativado.

## Arquivo .env Completo de Exemplo

Veja o arquivo `.env.example` para todas as variáveis de configuração disponíveis.
