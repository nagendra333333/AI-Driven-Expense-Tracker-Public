# AI-Powered Personal Expense Tracker

> Automate your monthly bank statement processing, expense categorization, and budget insights — end to end.

Built out of frustration with manually entering hundreds of transactions every month into the Money Manager app. This system reads your bank statements, cleans the data, categorizes every transaction using AI, exports to Money Manager format, and generates an AI-written budget insights report — all in under 2 minutes.

---

## About The Project

Every month I was spending 3–4 hours downloading bank statements, copying transactions, and categorizing them one by one. After completing an AI course covering open source models, RAG, and fine tuning — I decided to automate the entire workflow.

This project handles:

- **Multi-bank statement parsing** — Canara, HDFC, KVB, Federal Bank (CSV and XLS formats)
- **Zerodha tradebook integration** — Aggregates stock buy/sell trades by symbol and date
- **AI-powered categorization** — Hybrid approach: keyword rules handle ~90% instantly, NVIDIA GPT-4o 120B handles the rest via parallel API calls
- **RAG-inspired memory** — Reads previous months' categorized files to learn your patterns and improve accuracy over time
- **Bank-to-bank transfer detection** — Same-day same-amount debits and credits across different banks are automatically removed
- **Money Manager export** — Outputs a `.tsv` file ready to import directly into the Money Manager app
- **Budget insights report** — AI-generated Word document comparing actuals vs budget with category analysis, investment tracking, and suggestions

---

## Built With

| Tool | Purpose |
|---|---|
| Python 3.12 | Core language |
| pandas | Data parsing and transformation |
| openpyxl / xlrd / xlwt | Excel read/write |
| openai (NVIDIA API) | AI categorization and insights |
| python-dotenv | Secure API key management |
| python-docx | Word document generation |
| concurrent.futures | Parallel API calls |

**AI Stack:**
- NVIDIA Inference API (`openai/gpt-oss-120b`) — transaction categorization + budget insights
- Keyword rule engine — instant local classification (~90% hit rate)
- RAG-inspired history lookup — learns from previous months' data

---

## Project Structure

```
project/
├── bank_merger_monthly.py       # Step 1 — Parse, merge, categorize bank statements
├── Money_manager_format.py      # Step 2 — Convert to Money Manager TSV format
├── budget_insights.py           # Step 3 — Generate AI budget insights Word report
├── categories.json              # Canonical category alias map (edit to add new categories)
├── .env.example                 # API key template — copy to .env and fill in
├── .gitignore                   # Prevents sensitive files from being committed
│
├── INPUT_FOLDER/                # Drop your bank statement files here each month
│   ├── CANARA.CSV
│   ├── HDFC.xls
│   ├── KVB.csv
│   ├── FEDERAL.csv
│   └── tradebook-GW5901-EQ.csv  # Zerodha tradebook (optional)
│
├── OUTPUT_FOLDER/               # Full_statement_YYYY-MM-DD.xlsx saved here
├── Money_Manager_output/        # Money_Manager_YYYY-MM-DD.tsv saved here
├── Budget/                      # Monthly_Budget_2026.xlsx goes here
└── Insights/                    # Insights_YYYY-MM-DD.docx saved here
```

---

## Getting Started

### Prerequisites

- Python 3.10 or higher
- A NVIDIA API key — sign up at [build.nvidia.com](https://build.nvidia.com)
- Bank statements exported as CSV or XLS from your bank's net banking portal
- (Optional) Zerodha tradebook CSV from Console → Reports → Tradebook

### Installation

**1. Clone the repository**
```bash
git clone https://github.com/your_username/ai-expense-tracker.git
cd ai-expense-tracker
```

**2. Install dependencies**
```bash
pip install pandas openpyxl xlrd xlwt openai python-dotenv python-docx
```

**3. Set up your API key**
```bash
cp .env.example .env
```
Open `.env` and add your NVIDIA API key:
```
NVIDIA_API_KEY=your-nvidia-api-key-here
```

**4. Create the required folders**
```bash
mkdir INPUT_FOLDER OUTPUT_FOLDER Money_Manager_output Budget Insights
```

**5. Add your budget file**

Place your `Monthly_Budget_2026.xlsx` in the `Budget/` folder.
The file should have 4 sheets: `Income`, `Expenses`, `Investments`, `Summary`.

---

## Usage

### Step 1 — Drop your bank files into INPUT_FOLDER

Name your files with the bank name so the script detects them automatically:

```
INPUT_FOLDER/
├── CANARA.csv
├── HDFC.xls
├── KVB.csv
├── FEDERAL.csv
└── tradebook-GW5901-EQ.csv
```

Supported banks: Canara, HDFC, KVB, Federal Bank, Yes Bank, Indian Bank.
Any other bank with a standard CSV format will be detected via the generic column-sniff parser.

### Step 2 — Run the scripts in order

```bash
# Step 1: Parse, merge, clean, and categorize
python bank_merger_monthly.py

# Step 2: Convert to Money Manager format
python Money_manager_format.py

# Step 3: Generate budget insights report
python budget_insights.py
```

### What happens in each step

**`bank_merger_monthly.py`**
- Detects each bank from the filename
- Parses and merges all statements into one dataset
- Removes bank-to-bank transfers (same-day, same-amount across banks)
- Detects duplicate transactions and flags them
- Removes Zerodha debit rows (covered by tradebook) and keeps Zerodha credit rows (dividends)
- Appends Zerodha tradebook rows with buy/sell handling
- Runs a 3-pass categorization:
  - Pass 1: History lookup from previous months (RAG-inspired, zero API calls)
  - Pass 2: Keyword rules (~150 patterns, instant)
  - Pass 3: NVIDIA GPT-4o 120B via parallel API calls for remaining unknowns
- Saves `Full_statement_YYYY-MM-DD.xlsx` to OUTPUT_FOLDER with:
  - Transactions sheet (with dropdowns for manual category correction)
  - Bank Transfers sheet (removed transfers recorded for reference)

**`Money_manager_format.py`**
- Reads the latest `Full_statement_*.xlsx`
- Maps fields to exact Money Manager TSV format
- Handles expense rows (debit) and income rows (credit) separately
- Saves `Money_Manager_YYYY-MM-DD.tsv` ready to import into the app

**`budget_insights.py`**
- Reads the latest Money Manager TSV and your budget Excel file
- Compares actuals vs budget for each category
- Handles yearly-tracked categories (Cloth, Bike maintenance, Gas) by summing all months YTD
- Calls NVIDIA GPT-4o 120B to generate a full narrative insights report
- Saves `Insights_YYYY-MM-DD.docx` to the Insights folder

### Output Excel columns

| Column | Description |
|---|---|
| S.No | Row number |
| Date | Transaction date |
| Description | Bank narration / UPI description |
| Debit | Amount debited (red, number format) |
| Credit | Amount credited (green, number format) |
| Bank | Source bank name |
| Expenses Category | AI-assigned expense category (dropdown) |
| Expenses sub-category | Sub-category e.g. stock symbol, Mutual Fund (dropdown) |
| Income Category | AI-assigned income category for credit rows (dropdown) |
| Income sub-category | Income sub-category (dropdown) |

### Adding a new bank

Add one line to `BANK_NAME_MAP` in `bank_merger_monthly.py`:
```python
"axisbank": ("Axis Bank", "Unknown"),   # filename keyword → (display name, parser key)
```

Then name your file `AXISBANK.csv` and drop it in INPUT_FOLDER. The generic parser handles standard CSV formats automatically.

### Adding a new expense category

Add one line to `categories.json`:
```json
"New Category": ["new category", "alias one", "alias two"]
```

No code changes needed. Both the categorization engine and budget insights script pick it up automatically.

---

## Expense Categories

Categories match the Money Manager app exactly. Yearly-tracked categories (marked with *) are compared against full-year budget rather than monthly.

`Food` · `Rent` · `Baby` · `Family` · `Dth + Ott + Net + Mobile` · `Transportation` · `Cloth*` · `Fuel` · `Bike maintenance*` · `Gas*` · `Drink` · `Movie` · `Electricity` · `water` · `Ironing` · `Haircut` · `Investment` · `Other` · `Spouse` · `Pregnancy` · `Electronic Gadgets` · `Car maintenance` · `Tour` · `Marriage` · `lending` · `Home things` · `Gift` · `Mobile accessories` · `Medicine` · `Miscellaneous`

---

## How the AI Categorization Works

```
Transaction Description
        ↓
Pass 1: History Lookup (RAG-inspired)
  → Reads all previous Full_statement_*.xlsx files
  → Exact match → merchant name match → accept only if unambiguous
  → ~70-80% hit rate after first month
        ↓ (remaining unknowns)
Pass 2: Keyword Rules (instant, no API)
  → 150+ patterns across all categories
  → Handles Groww/SIP → Investment, TASMAC → Drink, etc.
  → ~90% overall hit rate
        ↓ (remaining ~5-10%)
Pass 3: NVIDIA GPT-4o 120B (parallel API calls)
  → Batches of 5 transactions
  → Up to 5 parallel threads
  → Strips UPI reference numbers before sending
  → Falls back to "Other" on API failure
```

---

## Security

- **Never commit your `.env` file** — it contains your API key
- The `.gitignore` already excludes `.env`, all statement folders, and Excel temp files
- Sample data in this repo uses a fictional person (`ARJUN KUMAR S`) with fabricated account numbers
- All real account numbers are masked as `XXXXXXXXXX`

---

## Roadmap

- [ ] Add `run_all.py` master script to run all 3 steps with one command
- [ ] Month-over-month trend comparison in budget insights report
- [ ] EMI and recurring payment auto-detection
- [ ] Duplicate transaction detection improvements (cross-bank external transfers)
- [ ] Sell trade profit/loss calculation in insights report
- [ ] Support for credit card statements
- [ ] Bar chart embed in Word insights report (matplotlib)

---

## Contributing

Contributions are welcome. If you want to add support for a new bank or improve categorization rules:

1. Fork the repo
2. Create a branch (`git checkout -b feature/add-axis-bank`)
3. Make your changes
4. Open a pull request with a description of what you changed and why

For bugs, open an issue with the bank name, a sanitized sample of the CSV (remove account numbers and personal details), and the error message.

---

## License

Distributed under the MIT License. See `LICENSE` for more information.

---

## Contact

Built by Nagendra Prasath M

LinkedIn — www.linkedin.com/in/nagendraprasath03121996
GitHub — [https://github.com/your_username](https://github.com/your_username)

---

## Acknowledgements

- [NVIDIA NIM](https://build.nvidia.com) — GPT-4o 120B inference API used for categorization and insights
- [Money Manager App](https://www.realbyteapps.com) — the app this system exports data into
- [pandas](https://pandas.pydata.org) — core data processing
- [openpyxl](https://openpyxl.readthedocs.io) — Excel generation and formatting
- [python-docx](https://python-docx.readthedocs.io) — Word document generation
