#!/usr/bin/env python3
import requests
import io
import os
from pathlib import Path

def test_simple_upload():
    print("=" * 70)
    print("TEST 1: Simple file upload (small file)")
    print("=" * 70)
    
    # Create a small test file
    test_file = io.BytesIO(b"Hello World - This is a test file content for upload testing!")
    
    files = {'file1': ('test_small.txt', test_file, 'text/plain')}
    data = {'action': 'simple'}
    
    try:
        response = requests.post('http://localhost:4050/tests/test_file_uploader_debug.asp', 
                                files=files, data=data, timeout=10)
        print(f"Status Code: {response.status_code}")
        print(f"Response Length: {len(response.text)} characters")
        
        # Check for success/error messages
        if "Upload Successful" in response.text:
            print("✓ SUCCESS: 'Upload Successful' message found")
            # Extract filename from response
            if "New Name:" in response.text:
                lines = response.text.split('\n')
                for i, line in enumerate(lines):
                    if "New Name:" in line:
                        print(f"  {line.strip()}")
        elif "Upload Failed" in response.text:
            print("✗ FAILED: 'Upload Failed' message found")
            lines = response.text.split('\n')
            for line in lines:
                if "Error:" in line:
                    print(f"  {line.strip()}")
        else:
            print("? No clear success/fail message found")
            # Show first 500 chars
            print("\nFirst part of response:")
            print(response.text[:500])
        
        # Check if file was actually created
        upload_dir = Path("www/uploads")
        if upload_dir.exists():
            files_in_dir = list(upload_dir.glob("*"))
            print(f"\nFiles in www/uploads: {len(files_in_dir)} files")
            if files_in_dir:
                # Show last 3 files
                recent = sorted(files_in_dir, key=lambda f: f.stat().st_mtime, reverse=True)[:3]
                for f in recent:
                    size = f.stat().st_size
                    print(f"  - {f.name} ({size} bytes)")
        else:
            print("✗ www/uploads directory not found!")
            
    except Exception as e:
        print(f"✗ Error: {type(e).__name__}: {e}")

def test_multiple_upload():
    print("\n" + "=" * 70)
    print("TEST 2: Multiple files upload")
    print("=" * 70)
    
    files = [
        ('files', ('file1.txt', io.BytesIO(b"Content file 1"), 'text/plain')),
        ('files', ('file2.txt', io.BytesIO(b"Content file 2 with more data here"), 'text/plain')),
        ('files', ('file3.jpg', io.BytesIO(b"Fake JPG content"), 'image/jpeg')),
    ]
    data = {'action': 'multiple'}
    
    try:
        response = requests.post('http://localhost:4050/tests/test_file_uploader_debug.asp',
                                files=files, data=data, timeout=10)
        print(f"Status Code: {response.status_code}")
        print(f"Response Length: {len(response.text)} characters")
        
        if "<table>" in response.text:
            print("✓ Table results found in response")
        else:
            print("? No table found in response")
        
        # Count OK/FAILED in response
        ok_count = response.text.count("<td class='success'>OK</td>")
        failed_count = response.text.count("<td class='error'>FAILED")
        print(f"  OK count: {ok_count}")
        print(f"  FAILED count: {failed_count}")
        
    except Exception as e:
        print(f"✗ Error: {type(e).__name__}: {e}")

def test_file_info():
    print("\n" + "=" * 70)
    print("TEST 3: File info (without upload)")
    print("=" * 70)
    
    test_file = io.BytesIO(b"This is content for file info test")
    files = {'infofile': ('info_test.pdf', test_file, 'application/pdf')}
    data = {'action': 'info'}
    
    try:
        response = requests.post('http://localhost:4050/tests/test_file_uploader_debug.asp',
                                files=files, data=data, timeout=10)
        print(f"Status Code: {response.status_code}")
        print(f"Response Length: {len(response.text)} characters")
        
        if "File Information" in response.text or "File Name:" in response.text:
            print("✓ File info found in response")
            # Extract info
            lines = response.text.split('\n')
            for line in lines:
                if "File Name:" in line or "Size:" in line or "MIME Type:" in line:
                    print(f"  {line.strip()}")
        else:
            print("? No file info found in response")
        
    except Exception as e:
        print(f"✗ Error: {type(e).__name__}: {e}")

if __name__ == '__main__':
    print("\nAxonASP File Uploader - Test Suite")
    print("Server: http://localhost:4050\n")
    
    test_simple_upload()
    test_multiple_upload()
    test_file_info()
    
    print("\n" + "=" * 70)
    print("Test suite completed!")
    print("=" * 70)
