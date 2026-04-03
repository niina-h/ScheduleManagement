"""本番サーバーでEdge headless PDF生成をテストするスクリプト"""
import subprocess, tempfile, os, sys, shutil

edge = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
print(f"Edge exists: {os.path.exists(edge)}")

with tempfile.NamedTemporaryFile(suffix=".html", delete=False, mode="w", encoding="utf-8") as f:
    f.write("<html><body><h1>Test PDF</h1></body></html>")
    tmp_html = f.name

tmp_pdf = tmp_html.replace(".html", ".pdf")
tmp_userdata = tempfile.mkdtemp(prefix="edge_test_")

cmd = [
    edge, "--headless", "--disable-gpu", "--no-sandbox",
    f"--user-data-dir={tmp_userdata}",
    f"--print-to-pdf={tmp_pdf}",
    "--no-pdf-header-footer",
    f"file:///{tmp_html.replace(os.sep, '/')}",
]
print(f"CMD: {cmd}")
print(f"tmp_html: {tmp_html}")
print(f"tmp_pdf: {tmp_pdf}")
print(f"TEMP dir: {tempfile.gettempdir()}")

try:
    r = subprocess.run(cmd, capture_output=True, timeout=30,
                       creationflags=subprocess.CREATE_NO_WINDOW)
    print(f"returncode: {r.returncode}")
    print(f"stdout: {r.stdout[:500]}")
    print(f"stderr: {r.stderr[:500]}")
    print(f"PDF exists: {os.path.exists(tmp_pdf)}")
    if os.path.exists(tmp_pdf):
        print(f"PDF size: {os.path.getsize(tmp_pdf)}")
except Exception as e:
    print(f"Error: {type(e).__name__}: {e}")
finally:
    for p in [tmp_html, tmp_pdf]:
        if os.path.exists(p):
            os.unlink(p)
    shutil.rmtree(tmp_userdata, ignore_errors=True)
