package com.houarizegai.webscraping.imdb;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;

import java.io.*;
import java.util.LinkedList;
import java.util.List;

public class    IMDBScraper {
    // Get the Top 250
    public static void main(String[] args) throws IOException {
        String url = "https://www.imdb.com/search/title/?groups=top_250&sort=user_rating";
        Document page = Jsoup.connect(url).get();
        Elements elementsContainer = page.getElementsByClass("lister-item-content");

        List<Film> topFilms = new LinkedList<>();

        elementsContainer.forEach(element -> {
            final String title = element.select("h3 > a").text();
            String year = element.getElementsByClass("lister-item-year").text();
            year = year.substring(1, year.length() - 1); // Remove () from the year
            final String category[] = element.select(".text-muted > .genre").text().split(", ");
            final String description = element.getElementsByTag("p").get(1).text();
            final String rate = element.getElementsByClass("ratings-imdb-rating").attr("data-value");

            topFilms.add(new Film(title, year, category, description, rate));

            // Writing to excel file

            Workbook workbook = new XSSFWorkbook();

            Sheet sheet = workbook.createSheet("Movies");
            sheet.setColumnWidth(0, 6000);
            sheet.setColumnWidth(1, 4000);

            Row header = sheet.createRow(0);

            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            XSSFFont font = ((XSSFWorkbook) workbook).createFont();
            font.setFontName("Title");
            font.setFontHeightInPoints((short) 16);
            font.setBold(true);
            headerStyle.setFont(font);

            Cell headerCell = header.createCell(0);
            headerCell.setCellValue("Title");
            headerCell.setCellStyle(headerStyle);

            headerCell = header.createCell(1);
            headerCell.setCellValue("Year");
            headerCell.setCellStyle(headerStyle);

            headerCell = header.createCell(2);
            headerCell.setCellValue("Category");
            headerCell.setCellStyle(headerStyle);

            headerCell = header.createCell(3);
            headerCell.setCellValue("Description");
            headerCell.setCellStyle(headerStyle);

            headerCell = header.createCell(4);
            headerCell.setCellValue("Rate");
            headerCell.setCellStyle(headerStyle);

            // cewll style
            CellStyle style = workbook.createCellStyle();
            style.setWrapText(true);

            Row row = sheet.createRow(topFilms.size());
            Cell cell = row.createCell(0);
            cell.setCellValue("John Smith");
            cell.setCellStyle(style);

            cell = row.createCell(1);
            cell.setCellValue(20);
            cell.setCellStyle(style);

            cell = row.createCell(2);
            cell.setCellValue(20);
            cell.setCellStyle(style);

            cell = row.createCell(3);
            cell.setCellValue(20);
            cell.setCellStyle(style);

            cell = row.createCell(4);
            cell.setCellValue(20);
            cell.setCellStyle(style);

            File currDir = new File(".");
            String path = currDir.getAbsolutePath();
            String fileLocation = path.substring(0, path.length() - 1) + "movies.xlsx";

            try {
                FileOutputStream outputStream = new FileOutputStream(fileLocation);
                workbook.write(outputStream);
                workbook.close();
            } catch (IOException e) {
                throw new RuntimeException(e);
            }


        });

        System.out.println(topFilms);
    }
}
