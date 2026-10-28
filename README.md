# Trading Company System

A command-line trading and inventory management system built with Python. The project uses **MySQL** for core application data and **MongoDB** for product reviews. It was developed as an object-oriented programming and database-learning project.

## Features

- User registration, login, password changes, and role-based menus.
- Separate roles for **User**, **Admin**, **Accountant**, and **CEO**.
- Product catalog and warehouse views for phones, cameras, and laptops.
- Product search and low-stock reporting.
- Buy and sell requests recorded as contracts.
- Admin approval or rejection of pending contracts.
- Wallet and company-balance tracking, including balance history.
- Profit/loss summaries and an Excel line chart.
- Product ratings and written reviews stored in MongoDB. A customer must have an approved purchase before reviewing a product, and a later review updates that customer's existing review for the product.
- Automatic MySQL table/database initialization and an optional one-time import from legacy JSON data when the MySQL users and products tables are empty.

## Technologies

- Python 3.10+
- MySQL
- MongoDB (reviews are optional; the main inventory and contract features can still run if MongoDB is unavailable)
- `mysql-connector-python`
- `pymongo`
- `openpyxl`
- `colorama`
- `python-dotenv`

## Requirements

Install the following before running the program:

1. Python 3.10 or newer.
2. A running MySQL server and a MySQL account allowed to create the project database and tables.
3. A running MongoDB server if you want product reviews enabled.

## Setup

### 1. Get the source code

Clone the repository or download its files. Open a terminal in the project directory.

### 2. Create and activate a virtual environment

**Windows PowerShell:**

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
```

If `py` is not available, try `python` instead.

### 3. Install Python dependencies

```bash
python -m pip install -r requirements.txt
```

### 4. Configure local environment variables

Copy the example configuration file to `.env`:

**Windows PowerShell:**

```powershell
Copy-Item .env.example .env
```

**macOS/Linux or Git Bash:**

```bash
cp .env.example .env
```

Open `.env` and set your own MySQL settings. Set `MYSQL_PASSWORD` to the password for your local MySQL account. Set `DEFAULT_CEO_PASSWORD` to a strong initial password for the CEO account. The example leaves passwords blank intentionally; configure them before the first run. Do not commit your real `.env` file to GitHub.

The default database name is `trading_system`. The program creates the database and required tables automatically if the configured MySQL account has sufficient privileges.

MongoDB defaults to `mongodb://localhost:27017/`. To use a different MongoDB instance, change `MONGODB_URI` in your local `.env` file. Keep any credentials inside that local file, not in source code.

### 5. Run the application

```bash
python trading_system.py
```

On the first run, the program initializes the database and creates a CEO account if one does not already exist and `DEFAULT_CEO_PASSWORD` is configured. Use `DEFAULT_CEO_PHONE` and `DEFAULT_CEO_PASSWORD` from your local `.env` file to log in. Change the initial credentials and use a non-root MySQL account for anything beyond local experimentation.

## Environment variables

| Variable | Purpose | Default |
|---|---|---|
| `MYSQL_HOST` | MySQL server host | `127.0.0.1` |
| `MYSQL_USER` | MySQL username | `root` |
| `MYSQL_PASSWORD` | MySQL password | Empty; configure locally |
| `MYSQL_DATABASE` | Database name | `trading_system` |
| `MONGODB_URI` | MongoDB connection URI | `mongodb://localhost:27017/` |
| `MONGODB_DATABASE` | MongoDB database used for reviews | `trading_system` |
| `DEFAULT_CEO_PHONE` | Initial CEO phone number | `09123456789` |
| `DEFAULT_CEO_PASSWORD` | Initial CEO password | Empty; configure locally |
| `DEFAULT_CEO_AGE` | Initial CEO age | `40` |

## Data and generated files

- The program creates and updates application data in MySQL and stores reviews in MongoDB.
- If legacy JSON files named `Users.json`, `Products.json`, `contracts.json`, `company_balance.json`, and `balance_history.json` exist and the MySQL users and products tables are empty, the program can import those files once. They are local data files and should not be uploaded if they contain personal or private information.
- The profit/loss chart is generated as `profit_loss_chart.xlsx` when the relevant report is requested. This generated file is excluded from Git by `.gitignore`.

## Security note

This is an educational project, not a production-ready trading platform. Review and improve authentication, password hashing, authorization, input validation, transaction handling, logging, and deployment configuration before using it with real users or financial data. Never commit database passwords, real `.env` files, personal data, or database exports.

## Project structure

```text
trading-company-system/
├── trading_system.py   # Application source code
├── README.md           # Project overview and setup instructions
├── requirements.txt    # Python dependencies
├── .env.example        # Safe configuration template (no real secrets)
└── .gitignore          # Excludes secrets, local data, and generated files
```
