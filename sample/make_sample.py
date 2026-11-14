"""Builds sample/network-tracker-sample.xlsx with made-up contacts, in the same layout as a real tracker."""
import datetime as dt
import os
import openpyxl

HEAD = ["Client", "Type", "Contact Name", "Contact Role", "Contact Number", "Contact Email", "Purpose",
        "Contact Method", "Target", "Prospect", "Meeting", "Contract", "Production", "Comments",
        "Last Contacted", "Follow-ups Sent", "Replied", "LinkedIn URL"]
ROWS = [
    ["Sunleaf Agro", "Exporter", "Tola Adeyemi", "CEO", None, None, "Sales", "LinkedIn", None, None, None, None, None, "Proposal sent 22/10", None, None, None, None],
    ["Harbour Capital", "PE Firm", "Ngozi Eze", "Partner", None, None, "Sales", "LinkedIn", None, None, None, None, None, "Proposal sent 20/10\nInitial comments: too expensive", None, None, None, None],
    ["Delta Polymers", "Importer", "Musa Bello", "MD", None, None, "Sales", "LinkedIn", None, None, None, None, None, None, dt.date(2026, 11, 3), 0, None, "https://www.linkedin.com/in/example-musa"],
    ["Crestline Machinery", "Supplier", "Ifeoma Obi", "Head of Sales", None, None, "Partnership", "LinkedIn", None, None, None, None, None, "First message 1/11", None, None, None, None],
    ["Greenfield Foods", "Manufacturer", "Kunle Ajayi", "Founder", None, None, "Network", "LinkedIn", None, None, None, None, None, "Last discussed 12/11", None, None, None, None],
    ["Northgate Partners", "PE Firm", "Aisha Musa", "Director", None, None, "Sales", "LinkedIn", None, None, None, None, None, None, None, None, None, None],
    ["Riverbend Oils", "Manufacturer", "Emeka Okoro", "COO", None, None, "Sales", "LinkedIn", None, None, None, None, None, "Partnership secured", None, None, None, None],
    ["Lagoon Logistics", "Supplier", "Bisi Cole", "CEO", None, None, "Partnership", "Referral", None, None, None, None, None, "Intro via Tola", None, None, None, None],
    ["Kola Ventures", "Individual", "Femi Lawal", "Investor", None, None, "Sales", "LinkedIn", None, None, None, None, None, None, dt.date(2026, 10, 20), 1, "Yes", None],
]

def build(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "KS"
    ws.append(HEAD)
    for r in ROWS:
        ws.append(r)
    wb.save(path)

if __name__ == "__main__":
    build(os.path.join(os.path.dirname(os.path.abspath(__file__)), "network-tracker-sample.xlsx"))
    print("Sample tracker written.")
