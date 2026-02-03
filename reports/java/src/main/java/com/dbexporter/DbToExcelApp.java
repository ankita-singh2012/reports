package com.dbexporter;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.*;
import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Database Query to Excel Export Application
 * Reads database configuration and query from properties file,
 * executes the query, and saves results to an Excel file.
 */
public class DbToExcelApp {
    
    private Properties config;
    private static final String CONFIG_FILE = "config.properties";
    
    public DbToExcelApp() throws IOException {
        this.config = loadConfig(CONFIG_FILE);
    }
    
    /**
     * Load configuration from properties file
     */
    private Properties loadConfig(String configFile) throws IOException {
        Properties props = new Properties();
        try (FileInputStream fis = new FileInputStream(configFile)) {
            props.load(fis);
            return props;
        } catch (IOException e) {
            throw new IOException("Config file not found: " + configFile, e);
        }
    }
    
    /**
     * Establish database connection
     */
    private Connection getConnection() throws SQLException {
        String url = String.format("jdbc:postgresql://%s:%s/%s",
            config.getProperty("db.host"),
            config.getProperty("db.port"),
            config.getProperty("db.name")
        );
        
        String user = config.getProperty("db.user");
        String password = config.getProperty("db.password");
        
        try {
            Class.forName("org.postgresql.Driver");
            Connection conn = DriverManager.getConnection(url, user, password);
            System.out.println("✓ Database connection established");
            return conn;
        } catch (ClassNotFoundException e) {
            throw new SQLException("PostgreSQL JDBC Driver not found", e);
        } catch (SQLException e) {
            System.out.println("✗ Connection failed: " + e.getMessage());
            throw e;
        }
    }
    
    /**
     * Execute SQL query and return result set metadata and data
     */
    private QueryResult executeQuery() throws SQLException {
        Connection conn = getConnection();
        try {
            String query = config.getProperty("query.sql");
            System.out.println("Executing query: " + query.substring(0, Math.min(50, query.length())) + "...");
            
            Statement stmt = conn.createStatement();
            ResultSet rs = stmt.executeQuery(query);
            
            QueryResult result = new QueryResult(rs);
            System.out.println("✓ Query executed successfully. Rows retrieved: " + result.getRowCount());
            
            return result;
        } finally {
            conn.close();
        }
    }
    
    /**
     * Save query results to Excel file
     */
    private void saveToExcel(QueryResult queryResult) throws IOException {
        String outputFile = config.getProperty("output.file");
        String sheetName = config.getProperty("output.sheet");
        
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet(sheetName);
            
            // Create header row
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < queryResult.getColumnCount(); i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(queryResult.getColumnName(i));
                cell.setCellStyle(getHeaderStyle(workbook));
            }
            
            // Create data rows
            List<Map<String, Object>> data = queryResult.getData();
            for (int rowIndex = 0; rowIndex < data.size(); rowIndex++) {
                Row dataRow = sheet.createRow(rowIndex + 1);
                Map<String, Object> rowData = data.get(rowIndex);
                
                for (int colIndex = 0; colIndex < queryResult.getColumnCount(); colIndex++) {
                    Cell cell = dataRow.createCell(colIndex);
                    Object value = rowData.get(queryResult.getColumnName(colIndex));
                    setCellValue(cell, value);
                }
            }
            
            // Auto-size columns
            for (int i = 0; i < queryResult.getColumnCount(); i++) {
                sheet.autoSizeColumn(i);
            }
            
            // Write to file
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                workbook.write(fos);
            }
            
            System.out.println("✓ Results saved to: " + outputFile);
            System.out.println("  Sheet name: " + sheetName);
            System.out.println("  Rows: " + data.size() + ", Columns: " + queryResult.getColumnCount());
            
        } catch (IOException e) {
            System.out.println("✗ Failed to save Excel file: " + e.getMessage());
            throw e;
        }
    }
    
    /**
     * Get header cell style
     */
    private CellStyle getHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.LIGHT_GRAY.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }
    
    /**
     * Set cell value based on object type
     */
    private void setCellValue(Cell cell, Object value) {
        if (value == null) {
            cell.setCellValue((String) null);
        } else if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
        } else {
            cell.setCellValue(value.toString());
        }
    }
    
    /**
     * Run the export process
     */
    public void run() {
        try {
            System.out.println("==================================================");
            System.out.println("Database Query to Excel Exporter (Java)");
            System.out.println("==================================================");
            
            QueryResult result = executeQuery();
            saveToExcel(result);
            
            System.out.println("==================================================");
            System.out.println("✓ Process completed successfully!");
            System.out.println("==================================================");
            
        } catch (Exception e) {
            System.err.println("\n✗ Process failed: " + e.getMessage());
            e.printStackTrace();
            System.exit(1);
        }
    }
    
    public static void main(String[] args) {
        try {
            DbToExcelApp app = new DbToExcelApp();
            app.run();
        } catch (IOException e) {
            System.err.println("Failed to load configuration: " + e.getMessage());
            e.printStackTrace();
            System.exit(1);
        }
    }
}
