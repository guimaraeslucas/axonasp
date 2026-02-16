<%@ Page Language="VBScript" %>
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>G3pix AxonASP REST API - Exemplos e Testes</title>
  <style>
    * { box-sizing: border-box; }
    body { 
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      margin: 0;
      padding: 20px;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      min-height: 100vh;
    }
    .container {
      max-width: 1200px;
      margin: 0 auto;
      background: white;
      border-radius: 12px;
      box-shadow: 0 20px 60px rgba(0,0,0,0.3);
      overflow: hidden;
    }
    header {
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      padding: 40px;
      text-align: center;
    }
    header h1 { margin: 0 0 10px 0; font-size: 28px; }
    header p { margin: 0; opacity: 0.9; }
    .content {
      padding: 40px;
    }
    .api-section {
      margin-bottom: 40px;
      border: 1px solid #e0e0e0;
      border-radius: 8px;
      padding: 20px;
      background: #f9f9f9;
    }
    .api-section h2 {
      margin-top: 0;
      color: #333;
      border-bottom: 2px solid #667eea;
      padding-bottom: 10px;
    }
    .endpoint {
      background: white;
      border-left: 4px solid #667eea;
      padding: 15px;
      margin: 15px 0;
      border-radius: 4px;
      font-family: 'Courier New', monospace;
      word-break: break-all;
    }
    .method {
      display: inline-block;
      padding: 4px 8px;
      border-radius: 4px;
      color: white;
      font-weight: bold;
      margin-right: 10px;
      min-width: 50px;
      text-align: center;
    }
    .method.get { background: #0066cc; }
    .method.post { background: #28a745; }
    .method.put { background: #ffc107; }
    .method.delete { background: #dc3545; }
    
    .description {
      margin: 10px 0;
      color: #666;
      font-family: Arial, sans-serif;
      font-style: italic;
    }
    .response-format {
      margin: 10px 0;
      padding: 10px;
      background: #f0f0f0;
      border-radius: 4px;
      font-size: 12px;
    }
    .example-code {
      background: #2d2d2d;
      color: #f8f8f2;
      padding: 15px;
      border-radius: 4px;
      margin: 10px 0;
      overflow-x: auto;
      font-family: 'Courier New', monospace;
      font-size: 12px;
      line-height: 1.5;
    }
    .try-btn {
      display: inline-block;
      padding: 10px 20px;
      background: #667eea;
      color: white;
      border: none;
      border-radius: 4px;
      cursor: pointer;
      text-decoration: none;
      font-size: 14px;
      margin: 10px 5px 10px 0;
      transition: background 0.3s;
    }
    .try-btn:hover {
      background: #5568d3;
    }
    .features {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
      gap: 20px;
      margin: 30px 0;
    }
    .feature-card {
      background: white;
      border: 1px solid #e0e0e0;
      border-radius: 8px;
      padding: 20px;
      text-align: center;
    }
    .feature-card h3 {
      margin-top: 0;
      color: #667eea;
    }
    .feature-card p {
      color: #666;
      font-size: 14px;
    }
    .format-selector {
      margin: 20px 0;
      padding: 15px;
      background: #f0f0f0;
      border-radius: 4px;
    }
    .format-selector label {
      margin-right: 20px;
    }
    .format-selector input {
      margin-right: 8px;
    }
  </style>
</head>
<body>
<div class="container">
  <header>
    <h1>🚀 G3pix AxonASP REST API</h1>
    <p>Exemplos práticos de uma API RESTful completa com GET, POST, PUT e DELETE</p>
  </header>

  <div class="content">
    
    <section class="api-section">
      <h2>📋 Visão Geral</h2>
      <p>Esta API demonstra o padrão REST implementado com AxonASP, utilizando as funções customizadas Ax para máxima compatibilidade e modernidade.</p>
      
      <div class="features">
        <div class="feature-card">
          <h3>🔄 Métodos HTTP</h3>
          <p>GET, POST, PUT, DELETE com roteamento automático</p>
        </div>
        <div class="feature-card">
          <h3>📄 Múltiplos Formatos</h3>
          <p>JSON (padrão) e HTML foram suportados</p>
        </div>
        <div class="feature-card">
          <h3>⚡ Funções Ax</h3>
          <p>Uso extensivo de funções customizadas modernas</p>
        </div>
        <div class="feature-card">
          <h3>🛡️ Validação</h3>
          <p>Validação de entrada e tratamento de erros</p>
        </div>
      </div>
    </section>

    <section class="api-section">
      <h2>🎯 Recursos Disponíveis</h2>

      <h3>Users (Usuários)</h3>
      
      <div class="endpoint">
        <span class="method get">GET</span>
        <span>/rest/users</span>
        <div class="description">Listar todos os usuários</div>
        <button class="try-btn" onclick="testEndpoint('/rest/users?route=users', 'GET')">Testar</button>
      </div>

      <div class="endpoint">
        <span class="method get">GET</span>
        <span>/rest/users/123</span>
        <div class="description">Obter usuário específico</div>
        <button class="try-btn" onclick="testEndpoint('/rest/users/123?route=users/123', 'GET')">Testar</button>
      </div>

      <div class="endpoint">
        <span class="method post">POST</span>
        <span>/rest/users</span>
        <div class="description">Criar novo usuário</div>
        <div class="example-code">
{
  "name": "New User",
  "email": "user@example.com",
  "role": "user"
}
        </div>
        <button class="try-btn" onclick="testEndpoint('/rest/users?route=users', 'POST')">Testar</button>
      </div>

      <div class="endpoint">
        <span class="method put">PUT</span>
        <span>/rest/users/123</span>
        <div class="description">Atualizar usuário existente</div>
        <div class="example-code">
{
  "name": "Updated Name",
  "email": "updated@example.com",
  "role": "admin"
}
        </div>
        <button class="try-btn" onclick="testEndpoint('/rest/users/123?route=users/123', 'PUT')">Testar</button>
      </div>

      <div class="endpoint">
        <span class="method delete">DELETE</span>
        <span>/rest/users/123</span>
        <div class="description">Deletar usuário</div>
        <button class="try-btn" onclick="testEndpoint('/rest/users/123?route=users/123', 'DELETE')">Testar</button>
      </div>

      <hr>

      <h3>Products (Produtos)</h3>
      
      <div class="endpoint">
        <span class="method get">GET</span>
        <span>/rest/products</span>
        <div class="description">Listar todos os produtos</div>
        <button class="try-btn" onclick="testEndpoint('/rest/products?route=products', 'GET')">Testar</button>
      </div>

      <div class="endpoint">
        <span class="method get">GET</span>
        <span>/rest/products/456</span>
        <div class="description">Obter produto específico</div>
        <button class="try-btn" onclick="testEndpoint('/rest/products/456?route=products/456', 'GET')">Testar</button>
      </div>

      <div class="endpoint">
        <span class="method post">POST</span>
        <span>/rest/products</span>
        <div class="description">Criar novo produto</div>
        <div class="example-code">
{
  "name": "New Product",
  "price": 99.99,
  "stock": 50
}
        </div>
        <button class="try-btn" onclick="testEndpoint('/rest/products?route=products', 'POST')">Testar</button>
      </div>

      <hr>

      <h3>Items (Itens com funções Ax avançadas)</h3>
      
      <div class="endpoint">
        <span class="method get">GET</span>
        <span>/rest/items</span>
        <div class="description">Listar itens (com AxExplode, AxWordCount, AxDate, AxTime)</div>
        <button class="try-btn" onclick="testEndpoint('/rest/items?route=items', 'GET')">Testar</button>
      </div>

      <div class="endpoint">
        <span class="method get">GET</span>
        <span>/rest/items/789</span>
        <div class="description">Obter item (com AxHash, AxBase64Encode, AxMd5, AxHtmlSpecialChars)</div>
        <button class="try-btn" onclick="testEndpoint('/rest/items/789?route=items/789', 'GET')">Testar</button>
      </div>

      <div class="endpoint">
        <span class="method post">POST</span>
        <span>/rest/items</span>
        <div class="description">Criar item (com AxUcFirst, AxCTypeAlpha)</div>
        <div class="example-code">
{
  "name": "New Item"
}
        </div>
        <button class="try-btn" onclick="testEndpoint('/rest/items?route=items', 'POST')">Testar</button>
      </div>

      <hr>

      <h3>Status (Health Check)</h3>
      
      <div class="endpoint">
        <span class="method get">GET</span>
        <span>/rest/status</span>
        <div class="description">Verificar status da API (com AxDate, AxTime, AxNumberFormat)</div>
        <button class="try-btn" onclick="testEndpoint('/rest/status?route=status', 'GET')">Testar</button>
      </div>

    </section>

    <section class="api-section">
      <h2>📝 Formato de Resposta</h2>
      
      <div class="format-selector">
        <label>
          <input type="radio" name="format" value="json" checked> 
          JSON (aplicação/json)
        </label>
        <label>
          <input type="radio" name="format" value="html"> 
          HTML Renderizado
        </label>
        <p style="color: #666; font-size: 12px; margin: 10px 0 0 0;">
          Adicione <code>?format=html</code> ou <code>&amp;format=html</code> à URL
        </p>
      </div>

      <h3>Estrutura de Resposta de Sucesso (200)</h3>
      <div class="example-code">
{
  "status": "success",
  "data": { ... },
  "count": 3,
  "timestamp": "2024-01-16T14:30:45"
}
      </div>

      <h3>Estrutura de Resposta de Erro (400, 404, 405, 422)</h3>
      <div class="example-code">
{
  "status": "error",
  "code": 404,
  "message": "Resource not found",
  "timestamp": "2024-01-16T14:30:45"
}
      </div>
    </section>

    <section class="api-section">
      <h2>🔧 Funções Ax Utilizadas</h2>
      <ul>
        <li><strong>AxExplode</strong> - Dividir string por delimitador</li>
        <li><strong>AxDate</strong> - Formatar data/hora</li>
        <li><strong>AxTime</strong> - Obter timestamp Unix</li>
        <li><strong>AxRand</strong> - Número aleatório</li>
        <li><strong>AxCount</strong> - Contar elementos de array</li>
        <li><strong>AxEmpty</strong> - Verificar se vazio</li>
        <li><strong>AxTrim</strong> - Remover espaços</li>
        <li><strong>AxUcFirst</strong> - Primeira letra maiúscula</li>
        <li><strong>AxWordCount</strong> - Contar palavras</li>
        <li><strong>AxHash / AxMd5</strong> - Hash de strings</li>
        <li><strong>AxBase64Encode / AxBase64Decode</strong> - Codificação Base64</li>
        <li><strong>AxHtmlSpecialChars</strong> - Escapar caracteres HTML</li>
        <li><strong>AxNumberFormat</strong> - Formatar números</li>
        <li><strong>AxCTypeAlpha</strong> - Validar caracteres alfabéticos</li>
      </ul>
    </section>

    <section class="api-section">
      <h2>💡 Exemplos com cURL</h2>
      
      <h3>GET - Listar Usuários</h3>
      <div class="example-code">
curl -X GET "http://localhost:4050/rest/users?route=users"
      </div>

      <h3>POST - Criar Usuário</h3>
      <div class="example-code">
curl -X POST "http://localhost:4050/rest/users?route=users" \
  -H "Content-Type: application/json" \
  -d "{\"name\": \"John Doe\", \"email\": \"john@example.com\", \"role\": \"user\"}"
      </div>

      <h3>PUT - Atualizar Usuário</h3>
      <div class="example-code">
curl -X PUT "http://localhost:4050/rest/users/1?route=users/1" \
  -H "Content-Type: application/json" \
  -d "{\"name\": \"Jane Doe\", \"email\": \"jane@example.com\", \"role\": \"admin\"}"
      </div>

      <h3>DELETE - Deletar Usuário</h3>
      <div class="example-code">
curl -X DELETE "http://localhost:4050/rest/users/1?route=users/1"
      </div>

      <h3>GET com HTML Format</h3>
      <div class="example-code">
curl -X GET "http://localhost:4050/rest/users?route=users&format=html"
      </div>
    </section>

    <section class="api-section">
      <h2>⚙️ Configuração web.config</h2>
      <p>As regras de rewrite foram adicionadas ao web.config para rotear URLs /rest/* para o front controller:</p>
      <div class="example-code">
&lt;rule name="RestAPI" stopProcessing="true"&gt;
  &lt;match url="^rest/(.*)$" /&gt;
  &lt;conditions&gt;
    &lt;add input="{REQUEST_FILENAME}" matchType="IsFile" negate="true" /&gt;
  &lt;/conditions&gt;
  &lt;action type="Rewrite" url="/rest/index.asp?route={R:1}" /&gt;
&lt;/rule&gt;
      </div>
    </section>

  </div>
</div>

<script>
  function testEndpoint(url, method) {
    const format = document.querySelector('input[name="format"]:checked').value;
    const fullUrl = url + (url.includes('?') ? '&' : '?') + 'format=' + format;
    
    window.open(fullUrl, '_blank');
  }
</script>

</body>
</html>
