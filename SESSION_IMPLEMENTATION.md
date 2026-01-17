# Sistema de Session Implementado - G3pix AxonASP

## Resumo

O sistema de Session foi implementado com sucesso para o interpretador ASP em Go. As sessões são agora armazenadas em arquivos locais na pasta `temp/session` em vez de apenas em memória.

## Características Implementadas

### 1. SessionManager (session_manager.go)
- **Gerenciamento de arquivo único de sessão**: Cada sessão é armazenada em um arquivo JSON separado (`temp/session/ASP<id>.json`)
- **Carregamento de sessão**: Carrega dados de sessão existente do arquivo
- **Salvamento de sessão**: Persiste dados de sessão para arquivo após cada requisição
- **Limpeza automática**: Remove sessões expiradas em background a cada 15 minutos
- **Timeout de sessão**: Padrão de 20 minutos (compatível com ASP clássico)

### 2. SessionObject (session_object.go)
- **Wrapper para dados de sessão**: Encapsula um mapa de dados com interface ASP-like
- **Acesso via índice**: Suporta `Session("chave")` para get e set
- **Acesso via propriedade**: Suporta `Session.SessionID` e outras propriedades
- **Propriedade SessionID**: Retorna o ID único da sessão

### 3. Integração com Executor (executor.go)
- **Carregamento de sessão**: Sessões são carregadas ao criar ExecutionContext
- **Salvamento de sessão**: Dados são persistidos em arquivo após execução do ASP
- **Cookie HTTP**: ASPSESSIONID é enviado no cookie HTTP automaticamente
- **Detecção de objetos built-in**: Session é reconhecido como objeto ASP built-in, não como variável

### 4. Inicialização do Servidor (main.go)
- **Cleanup automático**: Inicia rotina de background para limpar sessões expiradas
- **Intervalo de limpeza**: A cada 15 minutos

## Estrutura de Arquivo de Sessão

Cada arquivo de sessão (`temp/session/ASP<id>.json`) contém:

```json
{
  "id": "ASP1768693765685678300",
  "data": {
    "username": "John",
    "counter": 5
  },
  "created_at": "2026-01-17T20:49:25.6856783-03:00",
  "last_accessed": "2026-01-17T20:49:25.7249758-03:00",
  "timeout": 20
}
```

## Uso em ASP

### Definir variável de sessão
```asp
Session("username") = "John Doe"
Session("visit_count") = 42
```

### Obter variável de sessão
```asp
Dim user
user = Session("username")
Response.Write user
```

### Acessar ID da sessão
```asp
Response.Write Session.SessionID
```

## Como Funciona

1. **Primeira Requisição**:
   - Server gera novo ASPSESSIONID
   - Cria arquivo de sessão em `temp/session/`
   - Envia ASPSESSIONID como cookie HTTP

2. **Requisições Subsequentes**:
   - Browser envia ASPSESSIONID no cookie
   - Server carrega dados de sessão do arquivo
   - Todas as modificações em Session são persistidas em arquivo
   - Cookie ASPSESSIONID é reenviado

3. **Limpeza**:
   - Background routine verifica a cada 15 minutos
   - Remove arquivos de sessão expirados (> 20 minutos sem acesso)

## Compatibilidade

- ✅ Sintaxe ASP clássica: `Session("key")` e `Session.SessionID`
- ✅ Case-insensitivity (chaves armazenadas em lowercase interno)
- ✅ Timeout automático de 20 minutos
- ✅ Persistência em arquivo entre reinicializações do servidor
- ✅ Limpeza automática de sessões expiradas

## Testes

Página de teste: `/test_session.asp`
- Define múltiplas variáveis de sessão
- Exibe SessionID
- Demonstra persistência de dados

Exemplo de saída:
```
Stored in Session: John Doe
Visit Count: 1
Session ID: ASP1768693765685678300
Email: user@example.com
Age: 25
Active: True
```

## Notas Técnicas

- Sessões são armazenadas em JSON para facilitar debug e auditoria
- Mutex RWLock em SessionManager garante thread-safety
- SessionObject delega acesso a propriedades/índices para o mapa de dados
- Cookies são definidos com `MaxAge=20*60` (20 minutos) para sincronizar com timeout da sessão
- A limpeza de sessões expiradas é não-bloqueante e roda em background
