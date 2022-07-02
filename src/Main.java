import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.Scanner;

public class Main {
    static int id = 1;
    static String dbName = "";
    static final String startString = "INSERT INTO account (id, email, enabled, password, role_id, locale_tag) VALUES ";


    public static void main(String[] args) {

        try(FileWriter writer = new FileWriter("F:\\result.txt", false))
        {
            // запись всей строки
            String text = parse("F:\\user.xls");
            writer.write(text);
            writer.flush();
        }
        catch(IOException ex){
            System.out.println(ex.getMessage());
        }

    }

    public static String parse(String fileName) {

        //инициализируем потоки
        StringBuilder result = new StringBuilder(startString);
        InputStream inputStream = null;
        HSSFWorkbook workBook = null;
        try {
            inputStream = new FileInputStream(fileName);
            workBook = new HSSFWorkbook(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        //разбираем первый лист входного файла на объектную модель
        Sheet sheet = workBook.getSheetAt(0);
        Iterator<Row> it = sheet.iterator();
        //проходим по всему листу
        while (it.hasNext()) {
            result.append("(").append(id);

            Row row = it.next();
            Iterator<Cell> cells = row.iterator();
            while (cells.hasNext()) {

                Cell cell = cells.next();
                int cellType = cell.getCellType();
                //перебираем возможные типы ячеек
                switch (cellType) {
                    case Cell.CELL_TYPE_STRING:
                        String value = cell.getStringCellValue();
                        if(value.equals("null")|| value.equals("true")|| value.equals("false")){
                            result.append(", ");
                            result.append(cell.getStringCellValue());
                        } else {
                            result.append(", ");
                            result.append("'");
                            result.append(cell.getStringCellValue());
                            result.append("'");
                        }
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        result.append(", ");
                        Integer i = (int) Double.parseDouble(String.valueOf(cell.getNumericCellValue()));
                        result.append(i);
                        break;
                    case Cell.CELL_TYPE_FORMULA:
                        result.append(", ");
                        result.append((int)Double.parseDouble(String.valueOf(cell.getNumericCellValue())));
                        break;
                    default:
                        break;
                }
            }
            result.append("),").append("\n");
            id++;
        }
        result.setLength(result.length()-2);
        result.append(";");
        return result.toString();
    }
}
