# Quick Start - ASP Module

Comece a usar o módulo ASP em 5 minutos!

## 🚀 Instalação

O módulo já está integrado como `asp` na pasta `go-asp/asp`.

```bash
# Navegue até a pasta do projeto
cd go-asp
```

## 📖 Seu Primeiro Parse

### 1. Parse Simples

```go
package main

import (
	"fmt"
	"asp"
)

func main() {
	aspCode := `
<html>
<body>
	<% 
		Dim message
		message = "Hello ASP!"
		Response.Write(message)
	%>
</body>
</html>
`

	parser := asp.NewASPParser(aspCode)
	result, err := parser.Parse()
	
	if err != nil {
		fmt.Println("Erro:", err)
		return
	}
	
	fmt.Printf("Blocos encontrados: %d\n", len(result.Blocks))
	fmt.Printf("Programas VB: %d\n", len(result.VBPrograms))
}
```

### 2. Validação Rápida

```go
aspCode := `<% Dim x %><html>...</html>`

valid, errors := asp.ValidateASP(aspCode)
if !valid {
	for _, err := range errors {
		fmt.Println("Erro:", err)
	}
}
```

### 3. Extrair Componentes

```go
// Extrair apenas HTML
html := asp.ExtractHTMLOnly(aspCode)

// Extrair apenas VBScript
vb := asp.ExtractVBScriptOnly(aspCode)
```

## 🎯 Tarefas Comuns

### Analisar Arquivo ASP

```go
import "io/ioutil"

func analyzeASPFile(filename string) {
	content, _ := ioutil.ReadFile(filename)
	parser := asp.NewASPParser(string(content))
	result, _ := parser.Parse()
	
	fmt.Printf("%s: %d blocos\n", filename, len(result.Blocks))
}
```

### Processar Múltiplos Arquivos

```go
func processManyFiles(files []string) {
	for _, file := range files {
		valid, _ := asp.ValidateASP(readFile(file))
		status := "✓"
		if !valid {
			status = "✗"
		}
		fmt.Printf("%s %s\n", status, file)
	}
}
```

### Obter Informações de Erro

```go
parser := asp.NewASPParser(aspCode)
result, _ := parser.Parse()

if len(result.Errors) > 0 {
	for i, err := range result.Errors {
		fmt.Printf("Erro %d:\n", i)
		fmt.Printf("  Mensagem: %v\n", err)
	}
}
```

### Usar Objetos ASP

```go
ctx := asp.NewASPContext()

// Server
encoded, _ := ctx.Server.CallMethod("URLEncode", "hello world")

// Response
ctx.Response.CallMethod("Write", "Hello")
buffer := ctx.Response.GetBuffer()

// Session
ctx.Session.SetProperty("userid", 123)
fmt.Println(ctx.Session.GetProperty("userid"))
```

### Analisar Código Complexo

```go
analyzer := asp.NewASPCodeAnalyzer()
analysis := analyzer.Analyze(aspCode)

fmt.Printf("Blocos: %d\n", analysis["total_blocks"])
fmt.Printf("Complexidade: %s\n", analysis["complexity"])

if patterns, ok := analysis["patterns_detected"].([]string); ok {
	fmt.Println("Padrões:", patterns)
}
```

### Formatar Código

```go
formatter := asp.NewASPFormatter(4)
formatted := formatter.Format(aspCode)

fmt.Println(formatted)
```

## 💡 Exemplos Prontos

Veja exemplos completamente funcionais:

```go
// Simples
aspCode := `<% Dim x %>`

// Múltiplos blocos
aspCode := `<% Dim x %><html></html><% x = 5 %>`

// Com banco de dados (exemplo)
aspCode := `
<%
	Dim conn
	Set conn = CreateObject("ADODB.Connection")
%>
<html><body>Database</body></html>
<%
	conn.Close
%>
`

// Com formulário
aspCode := `
<%
	If Request.Form("submit") Then
		Dim user
		user = Request.Form("username")
		Response.Write("Welcome " & user)
	End If
%>
<form method="post">
	<input type="text" name="username">
	<input type="submit" name="submit" value="Go">
</form>
`
```

## 🔧 Integração com VBScript-Go

O módulo ASP reutiliza o parser VBScript-Go sem quebrar nada:

```go
import (
	"asp"
	vb "github.com/guimaraeslucas/vbscript-go"
)

func integrationExample(aspCode string) {
	// Parse como ASP
	aspParser := asp.NewASPParser(aspCode)
	aspResult, _ := aspParser.Parse()
	
	// Extrair VBScript puro
	vbCode := aspParser.ExtractVBScriptCode("\n")
	
	// Usar com VBScript-Go
	vbParser := vb.NewParser(vbCode)
	program := vbParser.Parse()
	
	fmt.Println("AST Program:", program)
}
```

## 📊 Performance

```go
import "time"

func benchmark() {
	aspCode := `<% Dim x %><html></html><% x = 5 %>`
	
	start := time.Now()
	for i := 0; i < 1000; i++ {
		parser := asp.NewASPParser(aspCode)
		parser.Parse()
	}
	elapsed := time.Since(start)
	
	fmt.Printf("1000 parses em: %v\n", elapsed)
}
```

## 🐛 Debug

```go
func debugASPCode(aspCode string) {
	parser := asp.NewASPParser(aspCode)
	result, err := parser.Parse()
	
	if err != nil {
		fmt.Println("❌ Erro crítico:", err)
		return
	}
	
	fmt.Println("✓ Parse bem-sucedido")
	
	for i, block := range result.Blocks {
		fmt.Printf("\nBloco %d:\n", i)
		fmt.Printf("  Tipo: %s\n", block.Type)
		fmt.Printf("  Linha: %d, Coluna: %d\n", block.Line, block.Column)
		fmt.Printf("  Conteúdo: %s\n", block.Content[:min(len(block.Content), 50)])
	}
	
	if len(result.Errors) > 0 {
		fmt.Println("\nErros de parse:")
		for _, e := range result.Errors {
			fmt.Printf("  ❌ %v\n", e)
		}
	}
}

func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}
```

## ⚡ Dicas de Performance

1. **Reutilize contextos:**
   ```go
   ctx := asp.NewASPContext()
   // Use ctx para múltiplas operações
   ```

2. **Valide antes de processar:**
   ```go
   valid, _ := asp.ValidateASP(code)
   if !valid { return }
   parser := asp.NewASPParser(code)
   ```

3. **Use batch processing:**
   ```go
   for _, file := range files {
       if valid, _ := asp.ValidateASP(file); valid {
           // Processe
       }
   }
   ```

## 🚨 Tratamento de Erros

```go
parser := asp.NewASPParser(aspCode)
result, err := parser.Parse()

// Erro crítico
if err != nil {
	log.Fatal("Parse error:", err)
}

// Erros de sintaxe
if len(result.Errors) > 0 {
	for _, parseErr := range result.Errors {
		fmt.Printf("Line %d: %v\n", parseErr.Line, parseErr)
	}
}

// Verificar blocos
if len(result.Blocks) == 0 {
	fmt.Println("Warning: No blocks found")
}
```

## 📚 Próximos Passos

1. ✅ Leia [README.md](README.md) para visão geral
2. ✅ Veja [examples.go](examples.go) para mais exemplos
3. ✅ Consulte [BEST_PRACTICES.md](BEST_PRACTICES.md) para boas práticas
4. ✅ Estude [STRUCTURE.md](STRUCTURE.md) para arquitetura
5. ✅ Rode testes: `go test ./asp`

## 🆘 Ajuda Rápida

```go
// Não sei por onde começar?
parser := asp.NewASPParser(aspCode)
result, _ := parser.Parse()
fmt.Printf("Encontrados %d blocos\n", len(result.Blocks))

// Quero apenas extrair código?
html := asp.ExtractHTMLOnly(aspCode)
vb := asp.ExtractVBScriptOnly(aspCode)

// Preciso validar?
valid, errors := asp.ValidateASP(aspCode)

// Quer análise profunda?
analyzer := asp.NewASPCodeAnalyzer()
analysis := analyzer.Analyze(aspCode)

// Necessita de objetos ASP?
ctx := asp.NewASPContext()
ctx.Response.CallMethod("Write", "Hello")
```

## ✅ Verificação

Você está pronto quando:
- [ ] Consegue fazer parse de código ASP simples
- [ ] Consegue extrair HTML e VBScript separadamente
- [ ] Consegue validar código ASP
- [ ] Entende como usar objetos ASP
- [ ] Consegue debugar erros

## 🎉 Parabéns!

Você está pronto para usar o módulo ASP. Para dúvidas:
- Consulte a documentação completa em [README.md](README.md)
- Veja exemplos em [examples.go](examples.go)
- Estude boas práticas em [BEST_PRACTICES.md](BEST_PRACTICES.md)

Divirta-se! 🚀
