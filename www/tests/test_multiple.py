import requests
import io

print("\n=== Test Upload Múltiplo ===\n")

# Multiple files test
files = [
    ('files', ('file1.txt', io.BytesIO(b"Content 1"), 'text/plain')),
    ('files', ('file2.txt', io.BytesIO(b"Content 2 longer"), 'text/plain')),
]
data = {"action": "multiple"}

response = requests.post("http://localhost:4050/tests/test_file_uploader_debug.asp", files=files, data=data)
print(f"Status: {response.status_code}")

# Extract table results
import re
table_match = re.search(r'<table>.*?</table>', response.text, re.DOTALL)
if table_match:
    table = table_match.group(0)
    # Find rows
    rows = re.findall(r'<tr>.*?</tr>', table, re.DOTALL)
    print(f"\nFound {len(rows)-1} files processed (header + data rows)")
    
    for i, row in enumerate(rows[1:]):  # Skip header
        # Extract columns
        cols = re.findall(r'<td[^>]*>(.*?)</td>', row, re.DOTALL)
        if len(cols) >= 5:
            print(f"\nFile {i+1}:")
            print(f"  Name: {cols[0].strip()}")
            print(f"  Size: {cols[1].strip()}")
            print(f"  MIME: {cols[2].strip()}")
            print(f"  Ext: {cols[3].strip()}")
            print(f"  Status: {cols[4].strip()}")
else:
    print("Tabela não encontrada na resposta")
    # Show debug
    if "DEBUG" in response.text:
        print("\nDebug messages found:")
        debug_lines = [line for line in response.text.split("\n") if "DEBUG" in line]
        for line in debug_lines[:5]:
            print(f"  {line}")
