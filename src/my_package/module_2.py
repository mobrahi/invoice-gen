from fpdf import FPDF
from datetime import datetime

class PDFInvoice(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 16)
        self.cell(0, 10, 'INVOICE', 0, 1, 'R')
        self.ln(10)
    
    def add_invoice(self, data):
        # Company info (left)
        self.set_font('Arial', '', 10)
        self.cell(100, 6, "Mohd Fairuz Bin Ibrahim", 0, 1)
        self.cell(100, 6, "ShopIQ S/B", 0, 1)
        self.cell(100, 6, "21st Floor, CMA Building", 0, 1)
        self.cell(100, 6, "64 Connaught Road Central", 0, 1)
        self.cell(100, 6, "Hong Kong", 0, 1)
        self.cell(100, 6, f"Email: fairuz@gospurr.com", 0, 1)
        self.ln(5)
        
        # Customer & Invoice details
        self.set_font('Arial', 'B', 10)
        self.cell(50, 8, f"Attn: {data['attn']}", 0, 1)
        self.cell(50, 8, f"Company: {data['company']}", 0, 1)
        self.cell(50, 8, f"Invoice No: {data['invoice_no']}", 0, 1)
        self.cell(50, 8, f"Date: {data['date'].strftime('%Y-%m-%d')}", 0, 1)
        self.ln(10)
        
        # Table header
        self.set_fill_color(230, 230, 230)
        self.set_font('Arial', 'B', 10)
        self.cell(80, 10, 'Description', 1, 0, 'C', 1)
        self.cell(30, 10, 'Unit Price', 1, 0, 'C', 1)
        self.cell(30, 10, 'Qty', 1, 0, 'C', 1)
        self.cell(40, 10, 'Total', 1, 1, 'C', 1)
        
        # Table rows
        self.set_font('Arial', '', 10)
        total = 0
        for item in data['items']:
            self.cell(80, 8, item['description'], 1)
            self.cell(30, 8, f"${item['unit_price']:,.2f}", 1, 0, 'R')
            self.cell(30, 8, str(item['qty']), 1, 0, 'C')
            item_total = item['unit_price'] * item['qty']
            self.cell(40, 8, f"${item_total:,.2f}", 1, 1, 'R')
            total += item_total
        
        # Grand Total
        self.set_font('Arial', 'B', 11)
        self.cell(140, 10, 'Grand Total (Nett):', 1, 0, 'R', 1)
        self.cell(40, 10, f"${total:,.2f}", 1, 1, 'R', 1)
        
        # Bank details
        self.ln(10)
        self.set_font('Arial', '', 9)
        self.multi_cell(0, 5, "Remarks:\n1. All prices are quoted in United States Dollars (USD)\n2. Please remit all funds via bank telegraphic transfer to:\n\nAccount Name: Mohd Fairuz Bin Ibrahim\nAccount No.: 121-255707-833\nBIC/SWIFT: HSBCHKHHHKH\nBank: HSBC Bank, Tsim Sha Tsui Premier Centre\n2/F & 3/F, HSBC Building Tsim Sha Tsui,\n82-84 Nathan Road, Tsim Sha Tsui, Kowloon Hong Kong")
        
        # Signature
        self.ln(10)
        self.cell(0, 6, "On Behalf of ShopIQ S/B", 0, 1)
        self.ln(5)
        self.cell(0, 6, "Mohd Fairuz Ibrahim", 0, 1)
        self.cell(0, 6, "Director", 0, 1)
        self.cell(0, 6, "[Electronic quotation, no signature required]", 0, 1)

# Usage
pdf = PDFInvoice()
pdf.add_page()
invoice_data = {
    'attn': 'David Chernitzky',
    'company': 'Armour Cyber LTD',
    'invoice_no': 'AC/002/2026',
    'date': datetime.now(),
    'items': [{'description': 'Delivery of services - December2025', 'unit_price': 7129.30, 'qty': 1}]
}
pdf.add_invoice(invoice_data)
pdf.output('Invoice_Generated.pdf')
print("✅ PDF invoice generated!")