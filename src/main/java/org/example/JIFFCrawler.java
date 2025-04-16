package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.FileOutputStream;
import java.io.IOException;

public class JIFFCrawler {
    public static void main(String[] args) {
        try {
            // Create a new Excel workbook
            Workbook workbook = new XSSFWorkbook();
            CreationHelper createHelper = workbook.getCreationHelper();

            // Create a new Excel sheet
            Sheet sheet = workbook.createSheet("Movie Schedule");

            for (int i = 0; i < 10; i++) {
                sheet.setColumnWidth(i, 256 * 26);
            }

            // Parse the HTML content
            // day
            int day = 6;
            Document doc = Jsoup.connect("https://www.jeonjufest.kr/Ticket/timetable_day.asp?dayNum=" + day).get();
            Elements theaterItems = doc.select(".schedule");
            System.out.println(theaterItems);
            Elements screenRounds = doc.select(".screen-round-wrap");
            System.out.println(screenRounds);

            CellStyle style = workbook.createCellStyle();
            style.setWrapText(true);
            style.setAlignment(HorizontalAlignment.CENTER);

            int rowIdx = 0;

            // Create a header row
            Row headerRow = sheet.createRow(rowIdx++);
            headerRow.createCell(0).setCellValue("극장");
            headerRow.getCell(0).setCellStyle(style);
            for (int i = 1; i <= 5; i++) {
                headerRow.createCell(i).setCellValue(i + "회");
                headerRow.getCell(i).setCellStyle(style);
            }

            // Iterate through the theater items

            for (Element round : screenRounds) {
                Element theaterNameEl = round.selectFirst(".thearter-name h3");
                if (theaterNameEl == null) continue;

                // Create a new row for each theater
                Row row = sheet.createRow(rowIdx++);
                row.setHeightInPoints(73);
                row.createCell(0).setCellValue(theaterNameEl.text());
                row.getCell(0).setCellStyle(style);

                // Get all screenings
                Elements screenings = round.select(".card-row .screen-sort:not(.empty)");
                for (Element screening : screenings) {
                    String sessionText = screening.select(".round-text").text(); // e.g., "3회"
                    if (sessionText.isEmpty()) continue;

                    int sessionNum = Integer.parseInt(sessionText.replaceAll("\\D+", ""));

//                    String category = screening.select(".category span").text();
                    Element titleElement = screening.selectFirst(".title a");
                    String title;
                    if (titleElement != null) {
                        title = titleElement.text();
                    } else {
                        title = screening.select(".title").text(); // fallback
                    }
                    String time = screening.select(".time .value").text();
                    String code = screening.select(".code .number").text();

                    StringBuilder movieDetails = new StringBuilder();
//                    .append(category).append("\n")
                    movieDetails
                            .append(code).append("\n")
                            .append(title).append("\n")
                            .append(time).append("\n→");

                    row.createCell(sessionNum).setCellValue(movieDetails.toString());
                    row.getCell(sessionNum).setCellStyle(style);
                }
            }


            // Write the Excel data to a file
            FileOutputStream fileOut = new FileOutputStream("movie_schedule_" + day + ".xlsx");
            workbook.write(fileOut);
            fileOut.close();

            // Close the workbook
            workbook.close();

            System.out.println("Excel file created successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
