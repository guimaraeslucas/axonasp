# Exemplos de uso do VBScript Parser

Este diretório contém exemplos de como usar o parser VBScript-Go.

## Como executar

### Exemplo 1: Executar com código inline
```bash
cd example
go run main.go
```

Este comando irá fazer o parsing dos exemplos de código VBScript embutidos no programa e exibir a estrutura AST.

### Exemplo 2: Executar com arquivo VBScript
```bash
cd example
go run main.go test1.vbs
```

ou

```bash
go run main.go test2.vbs
```

## Arquivos de exemplo

- **main.go**: Programa principal que demonstra o uso do parser
- **test1.vbs**: Exemplo com funções, variáveis e chamadas
- **test2.vbs**: Exemplo com estruturas de controle complexas (For, Select Case, Do While)

## Saída esperada

O programa irá exibir:
- ✅ Status do parsing (sucesso ou erro)
- 📋 Se Option Explicit está ativado
- 📊 Número de statements no programa
- 📝 Estrutura do AST com todos os nós

### Exemplo de saída

```
=== Exemplo 1: Código Inline ===
✅ Parsing bem-sucedido!
📋 Option Explicit: true
📊 Número de statements: 7

📝 Estrutura do AST:
1. VariablesDeclaration (3 variáveis)
   - x
   - y
   - z
2. AssignmentStatement: x = 10
3. AssignmentStatement: y = 20
4. AssignmentStatement: z = (x + y)
5. IfStatement
   Condition: (z > 25)
   Then (1 statements)
     1. CallStatement: Response.Write("Z é maior que 25")
   Else (1 statements)
     1. CallStatement: Response.Write("Z é menor ou igual a 25")
6. FunctionDeclaration: Soma (2 parâmetros, 1 statements)
7. SubDeclaration: ExibeMensagem (1 parâmetros, 1 statements)
```

## Como funciona

1. **Parser Creation**: `vbs.NewParser(code)` cria uma nova instância do parser
2. **Parsing**: `parser.Parse()` faz o parsing do código e retorna um AST
3. **AST Traversal**: O programa percorre o AST e exibe informações sobre cada nó
4. **Error Handling**: Usa `defer/recover` para capturar erros de parsing

## Tipos de nós AST suportados

- **VariablesDeclaration**: Declaração de variáveis (Dim)
- **AssignmentStatement**: Atribuição de valores
- **IfStatement**: Estrutura If/ElseIf/Else
- **ForStatement**: Loop For/Next
- **WhileStatement**: Loop While/Wend
- **DoStatement**: Loop Do/Loop
- **SelectStatement**: Select Case
- **FunctionDeclaration**: Declaração de função
- **SubDeclaration**: Declaração de sub-rotina
- **CallStatement**: Chamada de procedimento
- **ExpressionStatement**: Statement de expressão

## Expressões suportadas

- **IdentifierExpression**: Identificadores/variáveis
- **LiteralExpression**: Literais (números, strings, etc.)
- **BinaryExpression**: Operações binárias (+, -, *, /, =, <, >, etc.)
- **UnaryExpression**: Operações unárias (-, Not, etc.)
- **MemberExpression**: Acesso a membros (objeto.propriedade)
- **IndexOrCallExpression**: Indexação ou chamada de função

## Próximos passos

Você pode:
- Adicionar mais arquivos VBScript de teste
- Estender o programa para gerar diferentes saídas (JSON, XML, etc.)
- Implementar um visitor pattern para processar o AST
- Criar ferramentas de análise estática de código VBScript
- Implementar um transpiler VBScript → JavaScript/Python
