import requests
import io

print("\n=== Test Upload Simples ===\n")

# Small text file
file_data = io.BytesIO(b"Hello World - This is a test file!")
files = {"file1": ("myfile.txt", file_data, "text/plain")}
data = {"action": "simple"}

response = requests.post("http://localhost:4050/tests/test_upload_simples.asp", files=files, data=data)
print(f"Status: {response.status_code}")
print(f"Sucesso: {'SUCESSO' in response.text}")
print(f"Erro: {'ERRO' in response.text}")

# Show response snippet
if "SUCESSO" in response.text:
    print("\n✓ ARQUIVO ENVIADO COM SUCESSO!")
    lines = response.text.split("\n")
    for i, line in enumerate(lines):
        if "SUCESSO" in line:
            print(f"Resultado:\n{lines[i]}")
            for j in range(1, 5):
                if i+j < len(lines):
                    print(f"{lines[i+j]}")
elif "ERRO" in response.text:
    print("\n✗ ERRO NO UPLOAD:")
    lines = response.text.split("\n")
    for i, line in enumerate(lines):
        if "ERRO" in line:
            for j in range(i, min(i+5, len(lines))):
                print(f"{lines[j]}")
            break
else:
    print("\n? Nao foi encontrada mensagem de sucesso ou erro")
    print(f"Primeira parte da resposta ({len(response.text)} chars):\n")
    print(response.text[:600])
