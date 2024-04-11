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
            int day = 2;
            Document doc = Jsoup.connect("https://www.jeonjufest.kr/Ticket/timetable_day.asp?dayNum=" + day).get();
//            Elements theaterItems = doc.select(".schedule");
//            System.out.println(theaterItems);
            Elements movieTimetable = doc.select(".movie-timetable .timetable > div");
//            System.out.println(movieTimetable);

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

            int index = 0;
            // Iterate through the theater items
            Row row = sheet.createRow(1);

            for (Element section : movieTimetable) {
                if (section.hasClass("thearter-name")) {
                    // Start a new row for each theater
                    row = sheet.createRow(rowIdx++);
                    row.setHeightInPoints(73);
                    row.createCell(0).setCellValue(section.text());
                    row.getCell(0).setCellStyle(style);
                } else if (section.hasClass("card-row")){
                    Elements screenings = section.select(".screen-sort.swiper-slide:not(.empty)");
//                    System.out.println(screenings);
                    for (Element screening : screenings) {
                        String sessionNumberText = screening.select(".mobile-sort").text(); // E.g., "3회"
                        int sessionNumber = Integer.parseInt(sessionNumberText.replaceAll("\\D+", ""));
                        String code = screening.select(".category .number").text();
                        Element titleElement = screening.select(".title a").first(); // Get the <a> element within .title
                        if (titleElement == null) {
                            continue;
                        }
                        String title = titleElement.text();
                        // TODO: Add link to the title
//                        String link = titleElement.absUrl("href");
//                        System.out.println(link);
                        String time = screening.select(".time span").text();

                        StringBuilder moviesDetails = new StringBuilder();
                        moviesDetails.append(code).append("\n").append(title).append("\n").append(time).append("\n→ ");
                        row.createCell(sessionNumber).setCellValue(moviesDetails.toString());
                        row.getCell(sessionNumber).setCellStyle(style);
                    }
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
