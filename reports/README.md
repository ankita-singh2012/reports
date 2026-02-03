# Database Query to Excel Export

Dual implementation (Python & Java) for exporting database query results to Excel files.

## Project Structure

```
reports/
├── python/                  # Python implementation
│   ├── db_to_excel.py      # Main Python script
│   ├── config.properties   # Configuration file
│   ├── requirements.txt    # Python dependencies
│   └── README.md          # Python-specific instructions
│
├── java/                    # Java implementation
│   ├── pom.xml            # Maven configuration
│   ├── config.properties  # Configuration file
│   ├── src/               # Java source code
│   └── README.md          # Java-specific instructions
│
└── README.md              # This file
```

## Quick Start

### Python Version
```bash
cd python
pip install -r requirements.txt
# Edit config.properties with your database details
python db_to_excel.py
```

### Java Version
```bash
cd java
mvn clean package
# Edit config.properties with your database details
java -jar target/db-to-excel-1.0.0.jar
```

## Features (Both Implementations)
- ✓ Reads database configuration from properties file
- ✓ Executes configurable SQL queries
- ✓ Exports results to Excel format
- ✓ Handles PostgreSQL databases (easily extensible)
- ✓ Comprehensive error handling
- ✓ Clean, documented code

## Configuration

Both implementations use the same `config.properties` format:

```properties
# Database Connection
db.host=localhost
db.port=5432
db.name=mydb
db.user=postgres
db.password=password123
db.driver=postgresql

# Query Configuration
query.sql=SELECT * FROM your_table LIMIT 100

# Output Configuration
output.file=query_results.xlsx
output.sheet=Results
```

## Comparison

| Feature | Python | Java |
|---------|--------|------|
| Setup | Simple (pip install) | Requires Maven |
| Performance | Good | Excellent |
| Memory Usage | Moderate | Moderate |
| IDE Support | Basic | Excellent (IntelliJ, Eclipse) |
| Packaging | Script | Executable JAR |
| Dependencies | 3 packages | Maven-managed |

## Requirements

### Python
- Python 3.7+
- pandas, openpyxl, psycopg2

### Java
- Java 11+
- Maven 3.6+

## License
MIT
