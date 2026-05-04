# Invoice Generator - Professional Invoicing Tool

A feature-rich desktop application for creating, managing, and exporting professional invoices with multi-item support. Built with Python and Tkinter, this tool allows you to generate both Excel and PDF invoices instantly.

## ✨ Features

- **Dynamic Multi-Item Support** - Add/remove items dynamically with auto-calculation
- **Export Formats** - Generate invoices in Excel (.xlsx) and PDF formats
- **Real-time Calculations** - Automatic subtotal and grand total updates
- **Professional Layout** - Clean, organized interface with company and customer sections
- **Currency Support** - Multiple currencies (USD, EUR, GBP, HKD, SGD)
- **Bank Details Section** - Edit and include bank transfer information
- **Scrollable Items Area** - Handle unlimited invoice items
- **File Dialog Integration** - Choose where to save your invoices

## 📋 Prerequisites

Before running this application, ensure you have Python 3.7+ installed on your system.

### Required Libraries

```bash
pip install openpyxl fpdf
```

## 🚀 Installation

1. **Clone the repository**

```bash
git clone https://github.com/yourusername/invoice-generator.git
cd invoice-generator
```

2. **Install dependencies**

```bash
pip install -r requirements.txt
```

3. **Run the application**

```bash
python invoice_generator.py
```

## 📖 Usage Guide

### Creating an Invoice

1. **Fill in your company information** (Left side - FROM section)
   - Your company name and address
   - Contact email

2. **Enter customer details** (Right side - TO section)
   - Attention person name
   - Customer company name
   - Customer email address

3. **Set invoice details**
   - Invoice number (auto-generated with current month/year)
   - Invoice date
   - Select currency

4. **Add invoice items**
   - Click `+ Add New Item` to create rows
   - Enter description, unit price, and quantity
   - Total automatically calculates
   - Remove items with the `✖ Remove` button

5. **Review totals**
   - Subtotal updates automatically
   - Grand total reflects the final amount

6. **Edit bank details** (if needed)
   - Modify the bank transfer information section

7. **Export your invoice**
   - Click `Generate Excel` for editable spreadsheet format
   - Click `Generate PDF` for professional printable version

### Saving Invoices

- A file dialog will appear when generating invoices
- Choose your desired location and filename
 - Files are automatically named: `Invoice_[INVOICE_NO]_[DATE].extension`

## 🗂️ Project Structure

```
invoice-generator/
│
├── invoice_generator.py    # Main application file
├── requirements.txt         # Python dependencies
├── README.md               # Project documentation
│
└── examples/               # Sample generated invoices
    ├── sample_invoice.xlsx
    └── sample_invoice.pdf
```

## 💻 Technical Details

### Built With

- **Python 3.7+** - Core programming language
- **Tkinter** - GUI framework (built into Python)
- **openpyxl** - Excel file generation
- **fpdf** - PDF generation

### Key Functions

| Function | Description |
|----------|-------------|
| `add_item_row()` | Dynamically adds new item row to the invoice |
| `update_all_totals()` | Calculates subtotal and grand total in real-time |
| `generate_excel()` | Creates formatted Excel invoice with borders and styling |
| `generate_pdf()` | Generates professional PDF invoice |
| `get_items_list()` | Collects all items for export |

## 📸 Screenshots

### Main Application Window
```
┌─────────────────────────────────────────────────────────────┐
│ INVOICE GENERATOR                                    [X]    │
├─────────────────────────────────────────────────────────────┤
│ ┌─ Invoice Header Information ──────────────────────────┐  │
│ │ FROM                    │ TO                          │  │
│ │ [Your Company Name]     │ Attention: [_______]       │  │
│ │ [Your Address]          │ Company: [_______]         │  │
│ │ [Email]                 │ Email: [_______]           │  │
│ └─────────────────────────────────────────────────────────┘  │
│ ┌─ Invoice Details ─────────────────────────────────────┐  │
│ │ Invoice No: [AC/04/2026]  Date: [2026-04-28]         │  │
│ └─────────────────────────────────────────────────────────┘  │
│ ┌─ Invoice Items ───────────────────────────────────────┐  │
│ │ Description    │ Unit Price │ Qty │ Total    │ Remove │  │
│ │ [__________]   │ [____]     │ [_]  │ [0.00]   │ [✖]   │  │
│ │ [+ Add New Item]                                      │  │
│ └─────────────────────────────────────────────────────────┘  │
│ [Generate Excel] [Generate PDF] [Clear All]                │
└─────────────────────────────────────────────────────────────┘
```

### Sample Excel Output

The generated Excel file includes:
- Formatted headers and borders
- Company and customer information
- Itemized list with calculations
- Subtotals and grand totals
- Bank transfer details
- Signature section

### Sample PDF Output

Professional PDF invoice with:
- Clean layout
- Table formatting
- Currency symbols
- All terms and conditions
- Electronic signature line

## 🔧 Customization

### Adding New Currencies

Modify the currency combobox in `setup_gui()`:

```python
self.currency = ttk.Combobox(invoice_frame, 
                             values=["USD", "EUR", "GBP", "HKD", "SGD", "YOUR_CURRENCY"], 
                             width=10)
```

### Changing Invoice Number Format

Modify the default invoice number generation:

```python
self.invoice_no.insert(0, f"CUSTOM/{datetime.now().strftime('%Y')}/{datetime.now().strftime('%m')}")
```

### Adding Tax Calculation

Add this function to include tax:

```python
def calculate_tax(self, subtotal, tax_rate=0.10):
    tax = subtotal * tax_rate
    grand_total = subtotal + tax
    return tax, grand_total
```

## 🐛 Troubleshooting

### Common Issues

**Issue:** `ModuleNotFoundError: No module named 'openpyxl'`

**Solution:** Install the missing module
```bash
pip install openpyxl
```

**Issue:** PDF generation fails on Linux/Mac

**Solution:** FPDF works on all platforms. For advanced features, consider reportlab:
```bash
pip install reportlab
```

**Issue:** Tkinter not found

**Solution:** On Linux, install python-tk:
```bash
sudo apt-get install python3-tk
```

## 📝 Requirements File

Create a `requirements.txt` file:

```txt
openpyxl>=3.0.0
fpdf>=1.7.2
```

## 🤝 Contributing

Contributions are welcome! Here's how you can help:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

### Development Ideas

- [ ] Add database support for invoice history
- [ ] Email integration to send invoices directly
- [ ] Save customer templates for repeat clients
- [ ] Add logo upload functionality
- [ ] Multiple language support
- [ ] Dark mode theme
- [ ] Bulk invoice generation
- [ ] QR code for payment links

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 👥 Authors

**Mohd Fairuz Bin Ibrahim** - *Initial work* - [YourGitHub](https://github.com/mobrahi)

## 🙏 Acknowledgments

- Open source community for excellent libraries
- All contributors and users of this tool

## 📧 Contact

For support or inquiries:
- Email: faairuz@gmail.com

## 🔄 Version History

- **1.0** (2026-04-28)
  - Initial release
  - Basic invoice generation
  - Excel and PDF export
  - Multi-item support

---

**Made with ❤️ using Python**
