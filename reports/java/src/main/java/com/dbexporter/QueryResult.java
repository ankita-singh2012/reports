package com.dbexporter;

import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.util.*;

/**
 * Holds query results with metadata
 */
public class QueryResult {
    private List<Map<String, Object>> data;
    private List<String> columnNames;
    private int columnCount;
    private int rowCount;
    
    public QueryResult(ResultSet rs) throws SQLException {
        this.data = new ArrayList<>();
        this.columnNames = new ArrayList<>();
        
        ResultSetMetaData metaData = rs.getMetaData();
        this.columnCount = metaData.getColumnCount();
        
        // Extract column names
        for (int i = 1; i <= columnCount; i++) {
            columnNames.add(metaData.getColumnName(i));
        }
        
        // Extract data
        while (rs.next()) {
            Map<String, Object> row = new LinkedHashMap<>();
            for (int i = 1; i <= columnCount; i++) {
                row.put(metaData.getColumnName(i), rs.getObject(i));
            }
            data.add(row);
        }
        
        this.rowCount = data.size();
    }
    
    public List<Map<String, Object>> getData() {
        return data;
    }
    
    public List<String> getColumnNames() {
        return columnNames;
    }
    
    public int getColumnCount() {
        return columnCount;
    }
    
    public int getRowCount() {
        return rowCount;
    }
    
    public String getColumnName(int index) {
        return columnNames.get(index);
    }
}
