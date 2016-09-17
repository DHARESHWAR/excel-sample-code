import com.jcg.example.Student;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Created by ganeshdhareshwar on 15/03/16.
 */
public class excelArticle {
    public static final String FILE_PATH = "/Users/ganeshdhareshwar/Documents/excel_exportnew.xlsx";

    public void readExcel() {
        try {
            FileInputStream file = new FileInputStream(new File(FILE_PATH));
            XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(file);
            XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()){
                    Cell cell = cellIterator.next();
                    System.out.print(cell.getStringCellValue() + "          ");
                }
                System.out.println();
            }
        }catch (Exception e){

        }
        }

    public void createExcel(List<Student> studentList) {
        Workbook workbook = new XSSFWorkbook();
        Sheet studentsSheet = workbook.createSheet("Students");
        int rowIndex = 0;
        for(Student student : studentList){
            Row row = studentsSheet.createRow(rowIndex++);
            int cellIndex = 0;
            //first place in row is name
            row.createCell(cellIndex++).setCellValue(student.getName());
            //second place in row is marks in maths
            row.createCell(cellIndex++).setCellValue(student.getMaths());
            //third place in row is marks in Science
            row.createCell(cellIndex++).setCellValue(student.getScience());
            //fourth place in row is marks in English
            row.createCell(cellIndex++).setCellValue(student.getEnglish());
        }
        try {
            FileOutputStream fos = new FileOutputStream(FILE_PATH);
            workbook.write(fos);
            fos.close();

            System.out.println(FILE_PATH + " is successfully written");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        List<Student> studentList = new ArrayList<Student>();
        studentList.add(new Student("Xavier","87","67","89"));
        studentList.add(new Student("Radcliff","88","74","90"));
        studentList.add(new Student("Watson","83","90","91"));
        studentList.add(new Student("john","88","92","84"));
        studentList.add(new Student("Lawrance","69","73","85"));

        excelArticle excel = new excelArticle();
        excel.createExcel(studentList);
        excel.readExcel();
    }
    }



