import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbookType;
import yahoofinance.Stock;
import yahoofinance.YahooFinance;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * This program gathers data about the stocks and outputs them in excel file
 *
 * @author Dominika Bobik
 * Last Modified: 1/23/2021
 * <p>
 * Using:
 * https://financequotes-api.com/javadoc/yahoofinance/YahooFinance.html
 */

public class Searcher {

    /**
     * This method reads stock tickers from the text file and adds them to the array
     *
     * @param fileName name of the text file where stock tickers are
     * @return array containing tickers
     */
    public String[] getTickers(String fileName) {
        Scanner scanFile = null;
        File file;
        String[] tickers = null;
        try {
            file = new File(fileName);
            scanFile = new Scanner(file);
            int lineCounter = 0;
            while (scanFile.hasNextLine()) {
                String line = scanFile.nextLine();
                lineCounter += 1;
            }
            tickers = new String[lineCounter];
            scanFile.reset();
            scanFile = new Scanner(file);
            int count = 0;
            while (scanFile.hasNextLine()) {
                String ticker = scanFile.nextLine();
                tickers[count] = ticker;
                count += 1;
            }
        } catch (FileNotFoundException e) {
            System.out.println("Invalid file name");
        } finally {
            if (scanFile != null) {
                scanFile.close();
            }
        }
        return tickers;
    }

    public static void main(String[] args) throws IOException {
        Searcher searcher = new Searcher();
        int numberOfTickers = searcher.getTickers("Tickers.txt").length;
        String[] tickers;
        tickers = searcher.getTickers("Tickers.txt");

        //Create workbook and sheet
        XSSFWorkbook workbook = new XSSFWorkbook(XSSFWorkbookType.XLSX);
        XSSFSheet sheet1 = workbook.createSheet("Stats");

        //Header row
        Row header = sheet1.createRow(0);
        CellStyle cs = workbook.createCellStyle();
        cs.setWrapText(true);
        cs.setAlignment(HorizontalAlignment.CENTER);

        //Header cells
        Cell cell0 = header.createCell(0);
        cell0.setCellValue("Company Name");
        cell0.setCellStyle(cs);

        Cell cell1 = header.createCell(1);
        cell0.setCellStyle(cs);
        cell1.setCellValue("Ticker");

        Cell cell2 = header.createCell(2);
        cell2.setCellStyle(cs);
        cell2.setCellValue("Price");

        Cell cell3 = header.createCell(3);
        cell3.setCellStyle(cs);
        cell3.setCellValue("Dividend yield $");

        Cell cell4 = header.createCell(4);
        cell4.setCellStyle(cs);
        cell4.setCellValue("P/B");

        Cell cell5 = header.createCell(5);
        cell5.setCellStyle(cs);
        cell5.setCellValue("P/E");

        Cell cell6 = header.createCell(6);
        cell6.setCellStyle(cs);
        cell6.setCellValue("Market Cap");


        Cell cell7 = header.createCell(7);
        cell7.setCellStyle(cs);
        cell7.setCellValue("Return on dividend %");

        //Create data rows
        for (int i = 1; i < numberOfTickers; i++) {
            Row row = sheet1.createRow(i);
            String ticker = tickers[i];
            System.out.println(ticker);
            Stock stock = YahooFinance.get(ticker);

            //Stock data
            String companyName;
            double currentPrice;
            double annualDividend;
            double priceToBook;
            double priceToEarnings;
            double marketCap;
            try{
            companyName = stock.getName();
            currentPrice = stock.getQuote().getPrice().doubleValue();
            annualDividend = stock.getDividend().getAnnualYield().doubleValue();
            priceToBook = stock.getStats().getPriceBook().doubleValue();
            priceToEarnings = stock.getStats().getPe().doubleValue();
            marketCap = stock.getStats().getMarketCap().doubleValue();}
            catch (NullPointerException e){
                continue;
            }

            //There might not be data for the following
            double prevClosePrice;
            try {
                prevClosePrice = stock.getQuote().getPreviousClose().doubleValue();
            } catch (NullPointerException e) {
                prevClosePrice = 0;
            }

            //Create cells
            for (int j = 0; j < 8; j++) {
                Cell cell = row.createCell(j);
                if (j == 0) {
                    //Name of the company
                    cell.setCellValue(companyName);
                } else if (j == 1) {
                    //Ticker
                    cell.setCellValue(ticker);
                } else if (j == 2) {
                    //Current price of the stock
                    cell.setCellValue(currentPrice);
                } else if (j == 3) {
                    //Annual dividend
                    cell.setCellValue(annualDividend);
                } else if (j == 4) {
                    //Price to book ratio (<2)
                    cell.setCellValue(priceToBook);
                } else if (j == 5) {
                    //Price to earnings ratio(~16)
                    cell.setCellValue(priceToEarnings);
                } else if (j == 6) {
                    //Market Cap
                    cell.setCellValue(marketCap);
                } else if (j == 7) {
                    //Return on dividend
                    cell.setCellValue((annualDividend / prevClosePrice) * 100);
                }
            }
        }

        sheet1.autoSizeColumn(0);
        FileOutputStream outputStream = null;

        try {
            outputStream = new FileOutputStream("Stocks.xlsx");
            workbook.write(outputStream);
        } catch (FileNotFoundException e) {
            System.out.println("ERROR: Close the file!");
        } finally {
            if (outputStream != null){
            outputStream.close();}
            workbook.close();
        }
    }
}
