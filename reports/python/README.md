# Database Query to Excel Export (Python)

A Python script that executes database queries and exports results to Excel files.

## Features
- Reads database configuration from `config.properties`
- Executes configurable SQL queries
- Exports results directly to Excel format
- Supports PostgreSQL databases

## Setup

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Configure Database Connection
Edit `config.properties` and update the following:
```properties
db.host=your_db_host
db.port=5432
db.name=your_database_name
db.user=your_username
db.password=your_password
```

### 3. Set Your Query
Update the `query.sql` parameter in `config.properties`:
```properties
query.sql=SELECT * FROM your_table WHERE condition = 'value'
```

### 4. Run the Script
```bash
python db_to_excel.py
```

## Configuration Options

| Option | Description | Example |
|--------|-------------|---------|
| `db.host` | Database host address | localhost |
| `db.port` | Database port | 5432 |
| `db.name` | Database name | mydb |
| `db.user` | Database user | postgres |
| `db.password` | Database password | password123 |
| `query.sql` | SQL query to execute | SELECT * FROM table |
| `output.file` | Output Excel filename | results.xlsx |
| `output.sheet` | Excel sheet name | Results |
