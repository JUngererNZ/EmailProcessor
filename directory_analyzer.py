#!/usr/bin/env python3
"""
Directory Analyzer for FML Freight Solutions
Analyzes directory structure and generates JSON data + populated checklist.md
"""

import os
import json
import datetime
from pathlib import Path
import re
from collections import defaultdict

class DirectoryAnalyzer:
    def __init__(self, root_path):
        self.root_path = Path(root_path)
        self.files_data = []
        self.directories = set()
        self.file_types = defaultdict(int)
        self.analysis_results = {}
        
    def analyze_directory(self):
        """Main analysis function"""
        print("🔍 Starting directory analysis...")
        
        # Walk through all files and directories
        for root, dirs, files in os.walk(self.root_path):
            # Skip the root directory itself in directories list
            if Path(root) != self.root_path:
                self.directories.add(str(Path(root).relative_to(self.root_path)))
            
            for file in files:
                file_path = Path(root) / file
                try:
                    # Get file stats
                    stat = file_path.stat()
                    relative_path = str(file_path.relative_to(self.root_path))
                    
                    file_info = {
                        "path": relative_path,
                        "name": file,
                        "size_bytes": stat.st_size,
                        "modification_date": datetime.datetime.fromtimestamp(stat.st_mtime).isoformat(),
                        "extension": file_path.suffix.lower(),
                        "full_path": str(file_path)
                    }
                    
                    self.files_data.append(file_info)
                    self.file_types[file_path.suffix.lower()] += 1
                    
                except (PermissionError, OSError) as e:
                    print(f"⚠️  Cannot access {file_path}: {e}")
                    continue
        
        # Generate analysis results
        self.generate_analysis()
        
        return self.analysis_results
    
    def extract_reference_data(self):
        """Extract reference data from file names and content"""
        reference_data = {
            "file_ref": "",
            "client_ref": "",
            "consignee": "",
            "description": "",
            "pin_no": "",
            "serial_no": "",
            "engine_no": "",
            "vessel_voy": "",
            "bill_no": "",
            "eta": "",
            "container_no": "",
            "transporter": ""
        }
        
        # Analyze file names for patterns
        for file_info in self.files_data:
            name = file_info['name'].upper()
            
            # Extract PIN numbers (6-8 digits)
            pin_match = re.search(r'\b(\d{6,8})\b', name)
            if pin_match and not reference_data['pin_no']:
                reference_data['pin_no'] = pin_match.group(1)
            
            # Look for serial numbers (SERIAL, SN, etc.)
            serial_match = re.search(r'(?:SERIAL|SN)[\s_-]*([A-Z0-9-]+)', name, re.IGNORECASE)
            if serial_match and not reference_data['serial_no']:
                reference_data['serial_no'] = serial_match.group(1)
            
            # Look for engine numbers (ENGINE, ENG, etc.)
            engine_match = re.search(r'(?:ENGINE|ENG)[\s_-]*([A-Z0-9-]+)', name, re.IGNORECASE)
            if engine_match and not reference_data['engine_no']:
                reference_data['engine_no'] = engine_match.group(1)
            
            # Look for vessel/voyage info
            vessel_match = re.search(r'(?:VESSEL|VOY|VSL)[\s_-]*([A-Z0-9-]+)', name, re.IGNORECASE)
            if vessel_match and not reference_data['vessel_voy']:
                reference_data['vessel_voy'] = vessel_match.group(1)
            
            # Look for bill numbers
            bill_match = re.search(r'(?:BILL|BL|INV)[\s_-]*([A-Z0-9-]+)', name, re.IGNORECASE)
            if bill_match and not reference_data['bill_no']:
                reference_data['bill_no'] = bill_match.group(1)
            
            # Look for container numbers (4 letters + 7 digits pattern)
            container_match = re.search(r'\b([A-Z]{4}\d{7})\b', name)
            if container_match and not reference_data['container_no']:
                reference_data['container_no'] = container_match.group(1)
        
        return reference_data
    
    def analyze_billing_rates(self):
        """Analyze billing and rates information"""
        billing_rates = {
            "client_invoice": {"requested": False, "received": False},
            "tariff_code": "",
            "weight_cube": {"length": "", "breadth": "", "height": ""},
            "cod_bv_booked": {"requested": False, "date_booked": ""},
            "fix_number": "",
            "client_po": {"number": ""},
            "quotation_nr": {"created": False, "number": ""},
            "rate_sheet": {"applicable": False}
        }
        
        # Analyze file names for billing-related documents
        for file_info in self.files_data:
            name = file_info['name'].upper()
            
            # Check for invoice files
            if any(keyword in name for keyword in ['INVOICE', 'INV', 'BILL']):
                billing_rates['client_invoice']['received'] = True
            
            # Check for PO files
            if 'PO' in name or 'PURCHASE ORDER' in name:
                billing_rates['client_po']['number'] = "Found in documents"
                billing_rates['client_invoice']['requested'] = True
            
            # Check for quotation files
            if 'QUOTATION' in name or 'QUOTE' in name:
                billing_rates['quotation_nr']['created'] = True
                billing_rates['quotation_nr']['number'] = "Found in documents"
            
            # Check for rate sheets
            if 'RATE' in name and 'SHEET' in name:
                billing_rates['rate_sheet']['applicable'] = True
            
            # Check for booking documents
            if any(keyword in name for keyword in ['BOOKING', 'BV', 'COD']):
                billing_rates['cod_bv_booked']['requested'] = True
                billing_rates['cod_bv_booked']['date_booked'] = "Found in documents"
            
            # Check for dimension documents
            if any(keyword in name for keyword in ['DIMENSION', 'DIM', 'WEIGHT', 'CUBE']):
                billing_rates['weight_cube']['length'] = "Found in documents"
                billing_rates['weight_cube']['breadth'] = "Found in documents"
                billing_rates['weight_cube']['height'] = "Found in documents"
        
        return billing_rates
    
    def analyze_tasks_progress(self):
        """Analyze tasks and progress based on document presence"""
        tasks = {
            "logistics_instructions": {
                "transport_instruction": {"status": "⚪", "details": ""},
                "x_haul_instruction": {"status": "⚪", "details": ""},
                "landing_order": {"status": "⚪", "details": ""},
                "warehouse_notification": {"status": "⚪", "details": ""}
            },
            "documentation_customs": {
                "anf": {"status": "⚪", "details": ""},
                "customs_entry": {"status": "⚪", "details": ""},
                "cargo_dues": {"status": "⚪", "details": ""},
                "bv_feri_ad": {"status": "⚪", "details": ""},
                "pre_alert": {"status": "⚪", "details": ""}
            },
            "release_delivery": {
                "line_invoice": {"status": "⚪", "details": ""},
                "release_letter": {"status": "⚪", "details": ""},
                "release_mail": {"status": "⚪", "details": ""},
                "pod": {"status": "⚪", "details": ""}
            },
            "financials_finalization": {
                "final_docs": {"status": "⚪", "details": ""},
                "costing": {"status": "⚪", "details": ""},
                "invoicing": {"status": "⚪", "details": ""},
                "acquittal": {"status": "⚪", "details": ""}
            }
        }
        
        # Analyze file names to determine task completion
        file_names = [f['name'].upper() for f in self.files_data]
        
        # Logistics & Instructions
        if any('TRANSPORT' in name or 'INSTRUCTION' in name for name in file_names):
            tasks['logistics_instructions']['transport_instruction']['status'] = "🔵"
            tasks['logistics_instructions']['transport_instruction']['details'] = "Transport instructions found"
        
        if any('WAREHOUSE' in name for name in file_names):
            tasks['logistics_instructions']['warehouse_notification']['status'] = "🔵"
            tasks['logistics_instructions']['warehouse_notification']['details'] = "Warehouse notification found"
        
        # Documentation & Customs
        if any('BV' in name or 'FERI' in name or 'AD' in name for name in file_names):
            tasks['documentation_customs']['bv_feri_ad']['status'] = "🔵"
            tasks['documentation_customs']['bv_feri_ad']['details'] = "BV/FERI/AD documents found"
        
        if any('CUSTOMS' in name or 'WE' in name or 'RIT' in name for name in file_names):
            tasks['documentation_customs']['customs_entry']['status'] = "🔵"
            tasks['documentation_customs']['customs_entry']['details'] = "Customs entry documents found"
        
        # Release & Delivery
        if any('RELEASE' in name for name in file_names):
            tasks['release_delivery']['release_letter']['status'] = "🔵"
            tasks['release_delivery']['release_letter']['details'] = "Release letter found"
        
        if any('POD' in name or 'PROOF OF DELIVERY' in name for name in file_names):
            tasks['release_delivery']['pod']['status'] = "🔵"
            tasks['release_delivery']['pod']['details'] = "Proof of delivery found"
        
        # Financials & Finalization
        if any('INVOICE' in name or 'BILL' in name for name in file_names):
            tasks['financials_finalization']['invoicing']['status'] = "🔵"
            tasks['financials_finalization']['invoicing']['details'] = "Invoicing documents found"
        
        if any('COSTING' in name for name in file_names):
            tasks['financials_finalization']['costing']['status'] = "🔵"
            tasks['financials_finalization']['costing']['details'] = "Costing documents found"
        
        return tasks
    
    def generate_visuals_verification(self):
        """Generate visuals and verification status"""
        visuals = {
            "we_pics": False,
            "xe_pics": False,
            "all_pictures_uploaded": False
        }
        
        file_names = [f['name'].upper() for f in self.files_data]
        
        if any('WE' in name and ('PIC' in name or 'IMAGE' in name) for name in file_names):
            visuals['we_pics'] = True
        
        if any('XE' in name and ('PIC' in name or 'IMAGE' in name) for name in file_names):
            visuals['xe_pics'] = True
        
        visuals['all_pictures_uploaded'] = visuals['we_pics'] and visuals['xe_pics']
        
        return visuals
    
    def generate_analysis(self):
        """Generate comprehensive analysis results"""
        # Basic statistics
        total_files = len(self.files_data)
        total_size_bytes = sum(f['size_bytes'] for f in self.files_data)
        
        # Find largest files
        largest_files = sorted(self.files_data, key=lambda x: x['size_bytes'], reverse=True)[:5]
        
        # Date analysis
        if self.files_data:
            dates = [datetime.datetime.fromisoformat(f['modification_date']) for f in self.files_data]
            most_recent = max(dates).strftime('%Y-%m-%d')
            date_range = f"{min(dates).strftime('%Y-%m-%d')} to {max(dates).strftime('%Y-%m-%d')}"
        else:
            most_recent = "No files found"
            date_range = "No files found"
        
        # Generate the complete analysis
        self.analysis_results = {
            "directory": str(self.root_path),
            "scan_date": datetime.datetime.now().isoformat(),
            "statistics": {
                "total_files": total_files,
                "total_directories": len(self.directories),
                "total_size_bytes": total_size_bytes,
                "total_size_mb": round(total_size_bytes / (1024 * 1024), 2),
                "file_types": dict(self.file_types),
                "largest_files": [
                    {
                        "name": f['name'],
                        "size_mb": round(f['size_bytes'] / (1024 * 1024), 2),
                        "path": f['path']
                    } for f in largest_files
                ],
                "date_analysis": {
                    "most_recent": most_recent,
                    "date_range": date_range
                }
            },
            "core_reference": self.extract_reference_data(),
            "billing_rates": self.analyze_billing_rates(),
            "tasks_progress": self.analyze_tasks_progress(),
            "visuals_verification": self.generate_visuals_verification(),
            "file_status": {
                "status": "🔵 In Progress",
                "generated_date": datetime.datetime.now().strftime('%Y-%m-%d')
            }
        }
    
    def save_json_output(self, output_path):
        """Save analysis results to JSON file"""
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(self.analysis_results, f, indent=2, ensure_ascii=False)
        print(f"✅ JSON output saved to: {output_path}")
    
    def generate_checklist_content(self):
        """Generate checklist.md content based on analysis"""
        ref = self.analysis_results['core_reference']
        billing = self.analysis_results['billing_rates']
        tasks = self.analysis_results['tasks_progress']
        visuals = self.analysis_results['visuals_verification']
        status = self.analysis_results['file_status']
        
        checklist_content = f"""# 📂 MASTER FILE CHECKLIST

> [!INFO] **Core Reference Data**
> **File Ref:** {ref['file_ref']}<br>
> **Client Ref:** {ref['client_ref']}<br>
> **Consignee:** {ref['consignee']}<br>
> **Description:** {ref['description']}<br>
> **PIN No:** {ref['pin_no'] if ref['pin_no'] else '[ ] Not found'}<br>
> **Serial No:** {ref['serial_no'] if ref['serial_no'] else '[ ] Not found'}<br>
> **Engine No:** {ref['engine_no'] if ref['engine_no'] else '[ ] Not found'}<br>
> **Vessel / Voy:** {ref['vessel_voy'] if ref['vessel_voy'] else '[ ] Not found'}<br>
> **Bill No:** {ref['bill_no'] if ref['bill_no'] else '[ ] Not found'}<br>
> **ETA:** {ref['eta'] if ref['eta'] else '[ ] Not found'}<br>
> **Container #:** {ref['container_no'] if ref['container_no'] else '[ ] Not found'}<br>
> **Transporter:** {ref['transporter'] if ref['transporter'] else '[ ] Not found'}<br>

---

### 💰 Billing & Rates
- **Client Invoice:** {'[x]' if billing['client_invoice']['requested'] else '[ ]'} Requested | {'[x]' if billing['client_invoice']['received'] else '[ ]'} Received <br>
- **Tariff Code:** {'[x]' if billing['tariff_code'] else '[ ]'} {billing['tariff_code'] if billing['tariff_code'] else ''} <br>
- **Weight & Cube:** L {'[x]' if billing['weight_cube']['length'] else '[ ]'} {'| B' if billing['weight_cube']['breadth'] else ''} {'[x]' if billing['weight_cube']['breadth'] else '[ ]'} {'| H' if billing['weight_cube']['height'] else ''} {'[x]' if billing['weight_cube']['height'] else '[ ]'} <br>
- **COD / BV Booked:** {'[x]' if billing['cod_bv_booked']['requested'] else '[ ]'} Requested | {'[x]' if billing['cod_bv_booked']['date_booked'] else '[ ]'} {'Date booked: ' + billing['cod_bv_booked']['date_booked'] if billing['cod_bv_booked']['date_booked'] else ''}<br>
- **Fix Number:** {billing['fix_number'] if billing['fix_number'] else '[ ] Not found'} <br>
- **Client PO:** {'[x]' if billing['client_po']['number'] else '[ ]'} {'Number: ' + billing['client_po']['number'] if billing['client_po']['number'] else 'Number [ ]'} <br>
- **Quotation NR:** {'[x]' if billing['quotation_nr']['created'] else '[ ]'} Created | {'[x]' if billing['quotation_nr']['number'] else '[ ]'} {'Number: ' + billing['quotation_nr']['number'] if billing['quotation_nr']['number'] else 'Number [ ]'} <br>
- **Rate Sheet:** {'[x]' if billing['rate_sheet']['applicable'] else '[ ]'} Applicable <br>

---

### ✅ Tasks & Progress
*Status Key: (D) = Drafted | (X) = Completed*

**Logistics & Instructions**
- {'[x]' if tasks['logistics_instructions']['transport_instruction']['status'] == '🔵' else '[ ]'} (D) | ({'X' if tasks['logistics_instructions']['transport_instruction']['status'] == '🔵' else 'D'}) — **Transport Instruction**: {tasks['logistics_instructions']['transport_instruction']['details']}
- {'[x]' if tasks['logistics_instructions']['x_haul_instruction']['status'] == '🔵' else '[ ]'} (D) | ({'X' if tasks['logistics_instructions']['x_haul_instruction']['status'] == '🔵' else 'D'}) — **X-Haul Instruction**: {tasks['logistics_instructions']['x_haul_instruction']['details']}
- {'[x]' if tasks['logistics_instructions']['landing_order']['status'] == '🔵' else '[ ]'} (D) | ({'X' if tasks['logistics_instructions']['landing_order']['status'] == '🔵' else 'D'}) — **Landing Order / Navis Update**: {tasks['logistics_instructions']['landing_order']['details']}
- {'[x]' if tasks['logistics_instructions']['warehouse_notification']['status'] == '🔵' else '[ ]'} (D) | ({'X' if tasks['logistics_instructions']['warehouse_notification']['status'] == '🔵' else 'D'}) — **Warehouse Notification**: {tasks['logistics_instructions']['warehouse_notification']['details']}

**Documentation & Customs**
- {'[x]' if tasks['documentation_customs']['anf']['status'] == '🔵' else '[ ]'} (D) | ({'X' if tasks['documentation_customs']['anf']['status'] == '🔵' else 'D'}) — **ANF (Arrival Notification)**: {tasks['documentation_customs']['anf']['details']}
- {'[x]' if tasks['documentation_customs']['customs_entry']['status'] == '🔵' else '[ ]'} (D) | ({'X' if tasks['documentation_customs']['customs_entry']['status'] == '🔵' else 'D'}) — **Customs Entry (WE/RIT)**: {tasks['documentation_customs']['customs_entry']['details']}
- {'[x]' if tasks['documentation_customs']['cargo_dues']['status'] == '🔵' else '[ ]'} (D) | ({'X' if tasks['documentation_customs']['cargo_dues']['status'] == '🔵' else 'D'}) — **Cargo Dues**: {tasks['documentation_customs']['cargo_dues']['details']}
- {'[x]' if tasks['documentation_customs']['bv_feri_ad']['status'] == '🔵' else '[ ]'} (D) | ({'X' if tasks['documentation_customs']['bv_feri_ad']['status'] == '🔵' else 'D'}) — **BV / FERI / AD**: {tasks['documentation_customs']['bv_feri_ad']['details']}
- {'[x]' if tasks['documentation_customs']['pre_alert']['status'] == '🔵' else '[ ]'} (D) | ({'X' if tasks['documentation_customs']['pre_alert']['status'] == '🔵' else 'D'}) — **Pre-Alert Sent**: {tasks['documentation_customs']['pre_alert']['details']}

**Release & Delivery**
- {'[x]' if tasks['release_delivery']['line_invoice']['status'] == '🔵' else '[ ]'} (D) | ({'X' if tasks['release_delivery']['line_invoice']['status'] == '🔵' else 'D'}) — **Line Invoice**: {tasks['release_delivery']['line_invoice']['details']}
- {'[x]' if tasks['release_delivery']['release_letter']['status'] == '🔵' else '[ ]'} (D) | ({'X' if tasks['release_delivery']['release_letter']['status'] == '🔵' else 'D'}) — **Release Letter**: {tasks['release_delivery']['release_letter']['details']}
- {'[x]' if tasks['release_delivery']['release_mail']['status'] == '🔵' else '[ ]'} (D) | ({'X' if tasks['release_delivery']['release_mail']['status'] == '🔵' else 'D'}) — **Release Mail (S/Line / Aero)**: {tasks['release_delivery']['release_mail']['details']}
- {'[x]' if tasks['release_delivery']['pod']['status'] == '🔵' else '[ ]'} (D) | ({'X' if tasks['release_delivery']['pod']['status'] == '🔵' else 'D'}) — **POD (Proof of Delivery)**: {tasks['release_delivery']['pod']['details']}

**Financials & Finalization**
- {'[x]' if tasks['financials_finalization']['final_docs']['status'] == '🔵' else '[ ]'} (D) | ({'X' if tasks['financials_finalization']['final_docs']['status'] == '🔵' else 'D'}) — **Final Docs (SSC / PN / CDG)**: {tasks['financials_finalization']['final_docs']['details']}
- {'[x]' if tasks['financials_finalization']['costing']['status'] == '🔵' else '[ ]'} (D) | ({'X' if tasks['financials_finalization']['costing']['status'] == '🔵' else 'D'}) — **Costing (Landside & Delivery)**: {tasks['financials_finalization']['costing']['details']}
- {'[x]' if tasks['financials_finalization']['invoicing']['status'] == '🔵' else '[ ]'} (D) | ({'X' if tasks['financials_finalization']['invoicing']['status'] == '🔵' else 'D'}) — **Invoicing**: {tasks['financials_finalization']['invoicing']['details']}
- {'[x]' if tasks['financials_finalization']['acquittal']['status'] == '🔵' else '[ ]'} (D) | ({'X' if tasks['financials_finalization']['acquittal']['status'] == '🔵' else 'D'}) — **Acquittal / GRN / XE**: {tasks['financials_finalization']['acquittal']['details']}

---

### 📸 Visuals & Verification
- {'[x]' if visuals['we_pics'] else '[ ]'} **WE Pics**
- {'[x]' if visuals['xe_pics'] else '[ ]'} **XE Pics**
- {'[x]' if visuals['all_pictures_uploaded'] else '[ ]'} **All Pictures Uploaded**

---

### 📧 Email Summary & Actions
**1. Detailed Timeline**
* [Based on file modification dates: {self.analysis_results['statistics']['date_analysis']['date_range']}]

**2. Internal Pack Note**
* *Notes for Filing:* Analysis completed on {status['generated_date']}. Total files found: {self.analysis_results['statistics']['total_files']}. Key documents identified for reference data extraction.

**3. Trail Summary**
* File analysis completed with {self.analysis_results['statistics']['total_files']} files processed
* Core reference data extracted from document names
* Task progress status determined based on document presence
* Visual verification status updated based on image files found

**4. Operations Team Checklist**
* [x] Directory analysis completed
* [x] File metadata extracted
* [x] Reference data populated
* [x] Task status determined
* [ ] Manual verification of extracted data
* [ ] Update any missing reference information

---
**File Status:** {status['status']}
*Generated on: {status['generated_date']}*
"""
        
        return checklist_content
    
    def save_checklist(self, output_path):
        """Save checklist.md file"""
        content = self.generate_checklist_content()
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"✅ Checklist saved to: {output_path}")


def main():
    """Main execution function"""
    import sys
    
    # Check if a directory path is provided as command line argument
    if len(sys.argv) > 1:
        target_dir = sys.argv[1]
        print(f"📁 Using provided directory: {target_dir}")
    else:
        # Ask user for directory path
        print("📁 Directory Analyzer for FML Freight Solutions")
        print("=" * 50)
        target_dir = input("Please enter the directory path to analyze (or press Enter to use current directory): ").strip()
        
        # If user presses Enter without input, use current directory
        if not target_dir:
            target_dir = os.path.dirname(os.path.abspath(__file__))
            print(f"📁 Using current directory: {target_dir}")
    
    # Validate directory exists
    if not os.path.exists(target_dir):
        print(f"❌ Error: Directory '{target_dir}' does not exist.")
        return
    
    if not os.path.isdir(target_dir):
        print(f"❌ Error: '{target_dir}' is not a directory.")
        return
    
    # Initialize analyzer
    analyzer = DirectoryAnalyzer(target_dir)
    
    # Perform analysis
    results = analyzer.analyze_directory()
    
    # Save JSON output
    json_output_path = os.path.join(target_dir, 'directory_analysis.json')
    analyzer.save_json_output(json_output_path)
    
    # Save checklist
    checklist_path = os.path.join(target_dir, 'checklist.md')
    analyzer.save_checklist(checklist_path)
    
    # Print summary
    print(f"\n📊 Analysis Summary:")
    print(f"   Files processed: {results['statistics']['total_files']}")
    print(f"   Directories found: {results['statistics']['total_directories']}")
    print(f"   Total size: {results['statistics']['total_size_mb']} MB")
    print(f"   File types: {len(results['statistics']['file_types'])}")
    
    print(f"\n✅ Analysis complete! Files generated:")
    print(f"   📄 {json_output_path}")
    print(f"   📄 {checklist_path}")


if __name__ == "__main__":
    main()
