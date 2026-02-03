"""
Database Query to Excel Export Script
Reads database configuration and query from properties file,
executes the query, and saves results to an Excel file.
"""

import configparser
import pandas as pd
import psycopg2
import openpyxl
from pathlib import Path
import sys


class DatabaseQueryExporter:
    """Handles database connections, queries, and Excel export."""
    
    def __init__(self, config_file):
        """Initialize with configuration file path."""
        self.config = self._load_config(config_file)
        
    def _load_config(self, config_file):
        """Load configuration from properties file."""
        if not Path(config_file).exists():
            raise FileNotFoundError(f"Config file not found: {config_file}")
        
        config = configparser.ConfigParser()
        config.read(config_file)
        return config
    
    def _get_connection(self):
        """Establish database connection based on config."""
        try:
            connection = psycopg2.connect(
                host=self.config.get('DEFAULT', 'db.host'),
                port=self.config.get('DEFAULT', 'db.port'),
                database=self.config.get('DEFAULT', 'db.name'),
                user=self.config.get('DEFAULT', 'db.user'),
                password=self.config.get('DEFAULT', 'db.password')
            )
            print("✓ Database connection established")
            return connection
        except psycopg2.Error as e:
            print(f"✗ Connection failed: {e}")
            raise
    
    def execute_query(self):
        """Execute the query from config and return results as DataFrame."""
        connection = self._get_connection()
        try:
            query = self.config.get('DEFAULT', 'query.sql')
            print(f"Executing query: {query[:50]}...")
            
            df = pd.read_sql_query(query, connection)
            print(f"✓ Query executed successfully. Rows retrieved: {len(df)}")
            return df
        finally:
            connection.close()
    
    def save_to_excel(self, dataframe):
        """Save DataFrame to Excel file."""
        output_file = self.config.get('DEFAULT', 'output.file')
        sheet_name = self.config.get('DEFAULT', 'output.sheet')
        
        try:
            dataframe.to_excel(output_file, sheet_name=sheet_name, index=False)
            print(f"✓ Results saved to: {output_file}")
            print(f"  Sheet name: {sheet_name}")
            print(f"  Rows: {len(dataframe)}, Columns: {len(dataframe.columns)}")
        except Exception as e:
            print(f"✗ Failed to save Excel file: {e}")
            raise
    
    def run(self):
        """Execute full pipeline: connect, query, export."""
        try:
            print("=" * 50)
            print("Database Query to Excel Exporter")
            print("=" * 50)
            
            df = self.execute_query()
            self.save_to_excel(df)
            
            print("=" * 50)
            print("✓ Process completed successfully!")
            print("=" * 50)
            
        except Exception as e:
            print(f"\n✗ Process failed: {e}", file=sys.stderr)
            sys.exit(1)


def main():
    """Main entry point."""
    config_file = "config.properties"
    
    exporter = DatabaseQueryExporter(config_file)
    exporter.run()


if __name__ == "__main__":
    main()
