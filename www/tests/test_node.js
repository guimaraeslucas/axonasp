var crypto = require('crypto');

// Função que cria e exibe o hash de uma mensagem
function generateHash(message) {
  console.log(`\n[Executando...] Gerando hash para a mensagem: "${message}"`);
  
  // Cria um objeto hash usando o algoritmo SHA-256
  // Adiciona a mensagem e converte o resultado final em uma string hexadecimal
  const hash = crypto.createHash('sha256').update(message).digest('hex');
  
  console.log(`Hash (SHA-256): ${hash}`);
}

var myMessage = "Olá, Gemini! Este é um teste de criptografia.";
var delayEmMilissegundos = 3000; // 3 segundos

console.log("Iniciando o programa...");
console.log(`O hash será calculado em ${delayEmMilissegundos / 1000} segundos...\n`);

// O setTimeout agenda a execução da função para o futuro
setTimeout(function() {
  generateHash(myMessage);
}, delayEmMilissegundos);

// Esta linha será impressa imediatamente, antes do hash ser gerado,
// demonstrando o comportamento assíncrono do Node.js.
console.log("Aguardando o temporizador... (Esta linha executa antes do setTimeout terminar)");