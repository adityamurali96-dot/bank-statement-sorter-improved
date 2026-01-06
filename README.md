# Bank Statement Sorter - Web Service

Upload PDF (including scanned) and Excel bank statements for automatic AI-powered categorization.

## Features

- **OCR Support**: Processes scanned/image-based PDFs using Tesseract OCR
- **Multiple Input Formats**: PDF, XLS, XLSX, CSV
- **Smart Categorization**: 20+ transaction categories
- **Excel Output**: 3 sheets (Deposits, Withdrawals, Summary)
- **Railway Optimized**: Ready for one-click deployment

## Quick Deploy to Railway

### Option 1: Using Dockerfile (Recommended)
1. Connect your GitHub repository to Railway
2. Railway will automatically detect the Dockerfile
3. Deploy - OCR dependencies are included

### Option 2: Using Nixpacks
If Dockerfile doesn't work, Railway will fall back to nixpacks.toml which also includes OCR dependencies.

## Local Development

### Prerequisites
- Python 3.11+
- Tesseract OCR installed
- Java (for tabula-py PDF processing)

### macOS Setup
```bash
brew install tesseract
brew install openjdk
```

### Ubuntu/Debian Setup
```bash
sudo apt-get install tesseract-ocr tesseract-ocr-eng default-jre
```

### Run Locally
```bash
pip install -r requirements.txt
python app.py
```
Visit http://localhost:8080

## File Structure

```
bank-statement-sort/
├── app.py              # Main Flask application
├── templates/
│   └── index.html      # Web interface
├── requirements.txt    # Python dependencies
├── Dockerfile          # Docker build config (OCR included)
├── railway.toml        # Railway deployment config
├── nixpacks.toml       # Nixpacks config (fallback)
├── Procfile           # Process file
└── README.md
```

## Transaction Categories

### Deposits
- Salary
- DEP TFR HRMS (Mobile, Labour, Cleansing, Briefcase, Furniture, Utility, Pest)
- DEP BANKS PERFORMANCE PLI
- CDS BASED PLI PAID FOR THE FY
- CEMTEX DEP INTER CIRCLE SPORTS
- Other Deposits

### Withdrawals
- Bank INTEREST
- DIRECT DR
- Transfer to own A/c
- WDL TFR
- UPI Payment
- NEFT/RTGS
- Other Withdrawals

## Output Format

The Excel output contains 3 sheets:

1. **Deposits**: All credit transactions grouped by category with date/amount columns
2. **Withdrawals**: All debit transactions grouped by category with date/amount columns
3. **Summary**: Summary table with totals per category

## API Endpoints

- `GET /` - Web interface
- `POST /upload` - Upload file for processing
- `GET /download/<filename>` - Download processed Excel file
- `GET /health` - Health check endpoint

## Troubleshooting

### OCR Not Working
- Ensure Tesseract is installed: `tesseract --version`
- Check if lang pack is installed: `tesseract --list-langs`

### PDF Parsing Fails
- Ensure Java is installed for tabula-py: `java -version`
- Try converting PDF to images first if heavily corrupted

### Railway Deployment Issues
- Check build logs for missing dependencies
- Ensure Dockerfile is being used (check railway.toml)
- Increase timeout if processing large files

## Environment Variables

- `PORT` - Server port (default: 8080)
- `PYTHONUNBUFFERED` - Python output buffering (default: 1)

## License

MIT
