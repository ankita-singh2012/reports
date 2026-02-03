# Database Query to Excel Export (Java)

A Java application that executes database queries and exports results to Excel files.

## Features
- Reads database configuration from `config.properties`
- Executes configurable SQL queries
- Exports results directly to Excel format with formatting
- Supports PostgreSQL databases
- Built with Maven for easy dependency management

## Setup

### 1. Prerequisites
- Java 11 or higher
- Maven 3.6+

### 2. Build the Project
```bash
mvn clean package
```

This creates an executable JAR file: `target/db-to-excel-1.0.0.jar`

### 3. Configure Database Connection
Edit `config.properties` and update the following:
```properties
db.host=your_db_host
db.port=5432
db.name=your_database_name
db.user=your_username
db.password=your_password
```

### 4. Set Your Query
Update the `query.sql` parameter in `config.properties`:
```properties
query.sql=SELECT * FROM your_table WHERE condition = 'value'
```

### 5. Run the Application
```bash
java -jar target/db-to-excel-1.0.0.jar
```

Or directly with Maven:
```bash
mvn exec:java -Dexec.mainClass="com.dbexporter.DbToExcelApp"
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

## Project Structure
```
java/
├── pom.xml                           # Maven configuration
├── config.properties                 # Configuration file
├── README.md                         # This file
└── src/main/java/com/dbexporter/
    ├── DbToExcelApp.java            # Main application class
    └── QueryResult.java             # Result data holder
```

## Dependencies
- PostgreSQL JDBC Driver
- Apache POI (Excel support)
- SLF4J (Logging)

## Error Handling
The application includes error handling for:
- Missing configuration file
- Database connection failures
- Query execution errors
- File write failures
