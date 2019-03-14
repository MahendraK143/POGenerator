package com.test;

import java.io.File;
import java.io.FileInputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {
	private void readExcel() throws Exception{
		Class<?> clazz = null;
        Map<String, Class<?>> headerMapProperties = null;
        List<StringBuilder> valuesOfList = new LinkedList<>();
        List<Object> dynamicPojosList = new LinkedList<>();
		FileInputStream fileInputStream = new FileInputStream(new File("/Users/709809/Practice Code/CoolCode/test spreadsheet.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.iterator();
        
        while (rowIterator.hasNext())
        {
            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();        
            if(row.getRowNum() == 0){
            	headerMapProperties = cellIterationForHeader(cellIterator);
            	clazz = PojoGenerator.generate("com.maha.rana.dynamic.pojo$ExcelPojo", headerMapProperties);	
            }else{            	
            	valuesOfList.add(cellIterationForValues(cellIterator));
            }
        }
        
		//Setter value dynamic mapping   
        mapColumnValuesToPojo(clazz, headerMapProperties, valuesOfList, dynamicPojosList);
        
        //getter value mapping
        getValuesFromPojo(clazz, headerMapProperties, dynamicPojosList);
        
        System.out.println("No of columns in header:" +headerMapProperties.size());
        System.out.println("No of rows excluding header:" +valuesOfList.size()+"|"+valuesOfList);
	}

	private void getValuesFromPojo(Class<?> clazz, Map<String, Class<?>> headerMapProperties,
			List<Object> dynamicPojosList)
					throws IllegalAccessException, InvocationTargetException, NoSuchMethodException {
		for(Map.Entry<String, Class<?>> me: headerMapProperties.entrySet()){
        	
        		String getterName = "get" + me.getKey().substring(0, 1).toUpperCase()
        				+  me.getKey().substring(1);  
        		
        		for(Object pojoObject:dynamicPojosList){
        			String result = (String) clazz.getMethod(getterName).invoke(pojoObject);
        			System.out.println(getterName+": " + result);
        		}
        	
        }
	}

	private void mapColumnValuesToPojo(Class<?> clazz, Map<String, Class<?>> headerMapProperties,
			List<StringBuilder> valuesOfList, List<Object> dynamicPojosList) throws InstantiationException,
					IllegalAccessException, InvocationTargetException, NoSuchMethodException {
		Object obj;
		for(StringBuilder str:valuesOfList){
        	String[] parts = str.toString().split("\\|");
        	int count = 0;
        	obj = clazz.newInstance();
        	for(Map.Entry<String, Class<?>> me: headerMapProperties.entrySet()){
        		String setterName = "set" + me.getKey().substring(0, 1).toUpperCase()
        				+  me.getKey().substring(1);
        		
        		if(count < parts.length){
            		clazz.getMethod(setterName, String.class).invoke(obj, parts[count]);
                	count++;	
            	}
        	}
        	
        	dynamicPojosList.add(obj);
        }
	}

	private StringBuilder cellIterationForValues(Iterator<Cell> cellIterator) {
		StringBuilder cellValuesBinding = new StringBuilder();
		while (cellIterator.hasNext()) 
		{
		    Cell cell = cellIterator.next();
		    switch (cell.getCellType()) 
		    {
		        case Cell.CELL_TYPE_NUMERIC:
		            System.out.print(cell.getNumericCellValue() + "\t");
		            cellValuesBinding.append(cell.getNumericCellValue()+"|");
		            break;
		        case Cell.CELL_TYPE_STRING:
		            System.out.print(cell.getStringCellValue() + "\t");
		            cellValuesBinding.append(cell.getStringCellValue()+"|");
		            break;
		    }
		}
		System.out.print("\n");
		return cellValuesBinding;
	}

	private Map<String, Class<?>> cellIterationForHeader(Iterator<Cell> cellIterator) {
		Map<String, Class<?>> excelHeaderMap = new LinkedHashMap<>();
		while (cellIterator.hasNext()) 
		{
		    Cell cell = cellIterator.next();
		    switch (cell.getCellType()) 
		    {
		        case Cell.CELL_TYPE_STRING:
		            excelHeaderMap.put(cell.getStringCellValue(), String.class);
		            break;
		    }
		}
		return excelHeaderMap;
	}
	
	public static void main(String args[]) throws Exception{
		ReadExcel excel = new ReadExcel();
		excel.readExcel();
	}
}
