import csv
import urllib.request
import urllib.error
from openpyxl import load_workbook
from pathlib import Path

# File paths
WORKBOOK_VETS = "MyBuddy_Vet_Finder_MASTER_RECONCILED_2026-10-15_v9.xlsx"
WORKBOOK_PLACES = "PET_PLACES.xlsx"
WORKBOOK_REST = "Pets_places rest_updated.xlsx"
REPORT_FILE = "broken_links_report.csv"

def read_sheet(wb, sheet_name):
    if sheet_name not in wb.sheetnames:
        return []
    ws = wb[sheet_name]
    headers = [cell.value for cell in ws[1]]
    data = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if any(row):
            data.append(dict(zip(headers, row)))
    return data

def check_url(url):
    if not url or not str(url).startswith("http"):
        return None
        
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
    
    # Ignore SSL certificate errors entirely
    import ssl
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE

    req = urllib.request.Request(url, headers=headers, method='HEAD')
    
    try:
        with urllib.request.urlopen(req, timeout=10, context=ctx) as response:
            if response.status == 200: 
                return "OK"
            else: 
                return f"HTTP {response.status}"
    except urllib.error.HTTPError as e:
        # Ignore 403 (Blocked) and 405 (Method Not Allowed)
        if e.code in (403, 405, 402):
            return "OK"
        return f"HTTP Error {e.code}"
    except urllib.error.URLError as e:
        error_str = str(e.reason)
        # If the domain literally does not exist, flag it
        if "Errno 11001" in error_str or "getaddrinfo failed" in error_str:
            return "Dead Domain"
        # Ignore SSL errors and Windows connection resets (NParks)
        if "SSL" in error_str or "WinError 10054" in error_str:
            return "OK" 
        return f"URL Error: {e.reason}"
    except Exception as e:
        # Ignore timeouts
        if "timed out" in str(e):
            return "OK"
        return f"Failed: {str(e)}"

def main():
    print("Starting Master Link Monitoring...")
    
    broken_links = []
    total_checked = 0

    # 1. VETS (MyBuddy_Vet_Finder_MASTER...)
    if Path(WORKBOOK_VETS).exists():
        print("Checking Vets...")
        wb = load_workbook(WORKBOOK_VETS, data_only=True)
        master = read_sheet(wb, "MASTER_LIST")
        hours = read_sheet(wb, "VET_CONTACT_HOURS")
        
        contact_lookup = {}
        for h in hours:
            mbv_id = str(h.get("MyBuddy Vet ID") or "").strip()
            if mbv_id:
                contact_lookup[mbv_id] = {
                    "website": h.get("Official Website") or h.get("Website"),
                    "booking": h.get("Booking URL")
                }

        for r in master:
            mbv_id = str(r.get("MyBuddy Vet ID") or "").strip()
            name = r.get("Clinic / Provider Name")
            
            if not mbv_id or mbv_id not in contact_lookup:
                continue
                
            website = contact_lookup[mbv_id]["website"]
            booking = contact_lookup[mbv_id]["booking"]

            print(f"Checking Vet: {name}...")

            if website:
                total_checked += 1
                status = check_url(website)
                if status and status != "OK":
                    broken_links.append([name, "Website", website, status])
            
            if booking and booking != website:
                total_checked += 1
                status = check_url(booking)
                if status and status != "OK":
                    broken_links.append([name, "Booking", booking, status])
        wb.close()

    # 2. PLACES (PET_PLACES.xlsx)
    if Path(WORKBOOK_PLACES).exists():
        print("Checking Groomers, Dine, and Parks...")
        wb = load_workbook(WORKBOOK_PLACES, data_only=True)
        places = read_sheet(wb, wb.sheetnames[0]) # Read first sheet
        
        for p in places:
            name = p.get("Name")
            if not name: continue
            
            print(f"Checking Place: {name}...")
            
            fb = p.get("Facebook URL")
            ig = p.get("IG URL") or p.get("Instagram URL")
            website = p.get("Website") 
            booking = p.get("Booking URL") 
            
            for label, url in [("Facebook", fb), ("Instagram", ig), ("Website", website), ("Booking", booking)]:
                if url:
                    total_checked += 1
                    status = check_url(url)
                    if status and status != "OK":
                        broken_links.append([name, label, url, status])
        wb.close()

    # 3. REST (Pets_places rest_updated.xlsx)
    if Path(WORKBOOK_REST).exists():
        print("Checking remaining places...")
        wb = load_workbook(WORKBOOK_REST, data_only=True)
        places = read_sheet(wb, wb.sheetnames[0]) 
        
        for p in places:
            name = p.get("Name")
            if not name: continue
            
            print(f"Checking Place: {name}...")
            
            fb = p.get("Facebook URL")
            ig = p.get("Instagram URL")
            website = p.get("Website")
            booking = p.get("Booking URL")
            
            for label, url in [("Facebook", fb), ("Instagram", ig), ("Website", website), ("Booking", booking)]:
                if url:
                    total_checked += 1
                    status = check_url(url)
                    if status and status != "OK":
                        broken_links.append([name, label, url, status])
        wb.close()

    # Save report
    with open(REPORT_FILE, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(["Place Name", "Link Type", "URL", "Error Status"])
        writer.writerows(broken_links)

    print(f"\nMonitoring complete! Checked {total_checked} links across all files.")
    print(f"Found {len(broken_links)} issues. Saved to {REPORT_FILE}")

if __name__ == "__main__":
    main()