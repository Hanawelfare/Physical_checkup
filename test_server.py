import urllib.request
import sys

port = 8081
base_url = f"http://127.0.0.1:{port}"

files_to_test = [
    "/",
    "/index.html",
    "/styles.css",
    "/app.js"
]

print("Starting server verification tests...")

success = True
for path in files_to_test:
    url = base_url + path
    try:
        req = urllib.request.Request(url, method="GET")
        with urllib.request.urlopen(req, timeout=5) as response:
            status = response.status
            content_type = response.headers.get("Content-Type", "")
            content = response.read(100) # Read first 100 bytes
            
            print(f"GET {path} -> Status: {status}, Content-Type: {content_type}")
            if status != 200:
                print(f"ERROR: Expected status 200, got {status}")
                success = False
            if path == "/styles.css" and "text/css" not in content_type:
                print(f"ERROR: Expected CSS content-type, got {content_type}")
                success = False
            if path == "/app.js" and "javascript" not in content_type:
                print(f"ERROR: Expected JS content-type, got {content_type}")
                success = False
    except Exception as e:
        print(f"GET {path} failed: {e}")
        success = False

if success:
    print("\nSUCCESS: All files served correctly with proper Content-Type headers!")
    sys.exit(0)
else:
    print("\nFAILURE: Some files were not served correctly.")
    sys.exit(1)
