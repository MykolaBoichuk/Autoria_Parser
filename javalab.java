import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import java.io.FileOutputStream;
import java.io.IOException;

public class AutoRiaParserExcel {

    public static void main(String[] args) {
        try {
            Document doc = Jsoup.connect("https://auto.ria.com/").get();
            Elements cars = doc.select(".ticket-item");
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Cars");
            int rowNum = 0;
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue("Марка");
            row.createCell(1).setCellValue("Модель");
            row.createCell(2).setCellValue("Рік");
            row.createCell(3).setCellValue("Ціна");
          
            for (Element car : cars) {
              
                String brand = car.select(".brand").text();
                String model = car.select(".model").text();
                String year = car.select(".year").text();
                String price = car.select(".price").text();

                row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(brand);
                row.createCell(1).setCellValue(model);
                row.createCell(2).setCellValue(year);
                row.createCell(3).setCellValue(price);
            }

            FileOutputStream fileOut = new FileOutputStream("cars.xlsx");

            workbook.write(fileOut);
            fileOut.close();
            workbook.close();

            System.out.println("Дані успішно зібрані та збережені у файл cars.xlsx");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
