import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;


class Employee {
    private String name;
    private int age;
    private String location;

    public void Employee(String name, int age, String location) {
        this.name = name;
        this.age = age;
        this.location = location;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public String getLocation() {
        return location;
    }

    public void setLocation(String location) {
        this.location = location;
    }
}

public class readFromExcelFile {
    private static final String source = "/Users/mithoon.k/Downloads/excel.xlsx";
    private static final String dest = "/Users/mithoon.k/Downloads/output.xlsx";
    private static final String [] columns = {"Name", "Age", "Location"};
    private static List<Employee> employeeList = new ArrayList<Employee>();
    static {
        Employee employee = new Employee();
        employee.setName("Sachin");
        employee.setAge(30);
        employee.setLocation("Mumbai");
        Employee employee1 = new Employee();
        employee1.setName("Dhoni");
        employee1.setAge(30);
        employee1.setLocation("Ranchi");
        Employee employee2 = new Employee();
        employee2.setName("Yuvraj");
        employee2.setAge(30);
        employee2.setLocation("Chandigarh");
        employeeList.add(employee);
        employeeList.add(employee1);
        employeeList.add(employee2);
        System.out.println(employee1.getAge());
    }
    private static void readAndPrintExcelSheet() throws IOException, InvalidFormatException {
        Workbook workbook = WorkbookFactory.create(new File(source));
        //Printing the number of sheets in the workbook
        System.out.println("Total number of sheets in the workbook is " + workbook.getNumberOfSheets());
        //Printing the names of sheet in the workbook
        for(Sheet sheet : workbook) {
            System.out.println(sheet.getSheetName());
        }
        Sheet firstSheet = workbook.getSheetAt(0);
        DataFormatter dataFormatter = new DataFormatter();
        for (Row row: firstSheet) {
            for (Cell cell: row) {
                System.out.print(dataFormatter.formatCellValue(cell) + "  ");
            }
            System.out.println(" ");
        }
    }

    private static void writeDataToExcelFile() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Employee");
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.RED.getIndex());

        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        Row headerRow = sheet.createRow(0);

        for (int i= 0 ; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }
        int rowNum = 1;
        for (int i = 0; i <employeeList.size(); i++) {
            Row newRow = sheet.createRow(rowNum++);
            Cell cell1 = newRow.createCell(0);
            cell1.setCellValue(employeeList.get(i).getName());

            Cell cell2 = newRow.createCell(1);
            cell2.setCellValue(employeeList.get(i).getAge());

            Cell cell3 = newRow.createCell(2);
            cell3.setCellValue(employeeList.get(i).getLocation());
            System.out.println(employeeList.get(i).getAge());

        }

        for(int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }
        FileOutputStream fileOutputStream = new FileOutputStream("dest.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();

    }
    public static void main(String[] args) {
        try {
            readAndPrintExcelSheet();
        } catch (Exception e) {
            System.out.println("Following exception occured" + e);
        }
        try {
            writeDataToExcelFile();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
