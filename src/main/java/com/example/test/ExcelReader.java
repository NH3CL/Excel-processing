package com.example.test;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileSystemView;
import java.io.*;
import java.util.*;

import static javax.swing.SwingUtilities.updateComponentTreeUI;
import static org.apache.poi.ss.usermodel.CellType.*;

public class ExcelReader {
    public static String XLSX_FILE_PATH;
    public static List<Entry> entries;
    public static List<Entry> results;
    public static List<Entry> march;
    public static List<Entry> april;
    public static List<Entry> may;

    public static Map<String, Integer> marchPlayers;
    public static Map<String, Integer> aprilPlayers;
    public static Map<String, Integer> mayPlayers;

    public static Map<String, Integer> triplets;
    public static MyFrame frame;
    public static List<String> months;
    public static Map<String, String> monthMap;

    public static void main(String[] args) throws IOException, ClassNotFoundException, UnsupportedLookAndFeelException, InstantiationException, IllegalAccessException {

        fillMonths();

        frame = new MyFrame();
        UIManager.setLookAndFeel("com.sun.java.swing.plaf.windows.WindowsLookAndFeel");
        updateComponentTreeUI(frame);

        frame.println("--  Keresd ki a feldolgozandó excel fájlt (zárd be az EXCEL programot) --\n 
                      -- A formátum kötött:  \n
                      -- [ID]	[IHonap]	[Telefonszam]	[E-mail]	AP	BD	BI oszlopoknak kell szerepelnie az A1-es cellától kezdődően \n
                       ");

        JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
        int returnValue = jfc.showOpenDialog(null);
        if (returnValue == JFileChooser.APPROVE_OPTION) {
            XLSX_FILE_PATH = jfc.getSelectedFile().getPath();
            frame.println("--  Kiválasztott fájl: " + XLSX_FILE_PATH + " --\n");
        }

        frame.println("--  Informacio: Excel fájl első munkalapja fejléc nélkül --\n");
        fillEntries(XLSX_FILE_PATH);
        frame.println("\n--  Előfeltétel: 1. hónap - 2. hónap - 3. hónap --\n");
        splitEntries();
        frame.println("\n--  Számítás: Melyik hónapban hányszor szerepel egy telefonszám --\n");
        fillPlayers();
        frame.println("\n--  Számítás: Melyik telefonszám szerepel mindhárom hónapban --");
        frame.println("--            Mennyi volt a legkevesebb a három hónapban     --\n");
        findTrilpets();
        frame.println("\n-- Eredmény: Egy jatékos összes adata arról a hónapról --");
        frame.println("--           amikor a legkevesebbet játszott           --\n");
        writeTriplets();

        writeSheet(XLSX_FILE_PATH);
        frame.println("\n-- Az eredmény megtalálhato a bemeneti excel fájl mellett EREDMENY.xls néven  --\n");

    }

    private static void writeSheet(String fileIn) throws IOException {
        Workbook workbook = new XSSFWorkbook();

        CreationHelper createHelper = workbook.getCreationHelper();
        Sheet sheet = workbook.createSheet("EREDMENY");

        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 12);

        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);
        String[] columns = {"[ID]", "[IHonap]", "[Telefonszam]", "[E-mail]", "AP", "BD", "BI"};

        int rowNum = 0;
        Row headerRow = sheet.createRow(rowNum++);
        // Create cells
        for(int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }

        for(Entry entry: results) {
            Row row = sheet.createRow(rowNum++);

            row.createCell(0)
                    .setCellValue(entry.getId());

            row.createCell(1)
                    .setCellValue(entry.getMonth());

            row.createCell(2)
                    .setCellValue(entry.getPhone());

            row.createCell(3)
                    .setCellValue(entry.getMail());

            row.createCell(4)
                    .setCellValue(entry.getAP());

            row.createCell(5)
                    .setCellValue(entry.getBN());

            row.createCell(6)
                    .setCellValue(entry.getBI());

        }

        // Resize all columns to fit the content size
        for(int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream(XLSX_FILE_PATH.substring(0, XLSX_FILE_PATH.lastIndexOf("\\")) + "\\EREDMENY.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();

    }

    private static void writeTriplets() {
        results = new ArrayList<Entry>();

        for(String phone: triplets.keySet()) {
            if(marchPlayers.get(phone) == triplets.get(phone)) {
                for(Entry entry : march) {
                    if(entry.getPhone().equals(phone)) {
                        frame.println(entry.toString());
                        results.add(entry);
                    }
                }
            } else if(aprilPlayers.get(phone) == triplets.get(phone)) {
                for(Entry entry : april) {
                    if(entry.getPhone().equals(phone)) {
                        frame.println(entry.toString());
                        results.add(entry);
                    }
                }
            }else if(mayPlayers.get(phone) == triplets.get(phone)) {
                for(Entry entry : may) {
                    if(entry.getPhone().equals(phone)) {
                        frame.println(entry.toString());
                        results.add(entry);
                    }
                }
            }
        }
    }

    private static void findTrilpets() {
        triplets = new HashMap<String, Integer>();

        for(String phone : marchPlayers.keySet()) {
            if(aprilPlayers.containsKey(phone) && mayPlayers.containsKey(phone)) {
                int minOccur = marchPlayers.get(phone);
                if(aprilPlayers.get(phone) < minOccur)
                    minOccur = aprilPlayers.get(phone);
                if(mayPlayers.get(phone) < minOccur)
                    minOccur = mayPlayers.get(phone);

                frame.println(phone + " with minimum plays: " + minOccur);
                triplets.put(phone, minOccur);
            }
        }
    }

    private static void fillPlayers() {
        marchPlayers = new HashMap<String, Integer>();
        aprilPlayers = new HashMap<String, Integer>();
        mayPlayers   = new HashMap<String, Integer>();

        for(Entry entry : march) {
            if(marchPlayers.containsKey(entry.getPhone())) {
                int currCount = marchPlayers.get(entry.getPhone());
                marchPlayers.put(entry.getPhone(), ++currCount);
            } else {
                marchPlayers.put(entry.getPhone(), 1);
            }
        }

        for(Entry entry : april) {
            if(aprilPlayers.containsKey(entry.getPhone())) {
                int currCount = aprilPlayers.get(entry.getPhone());
                aprilPlayers.put(entry.getPhone(), ++currCount);
            } else {
                aprilPlayers.put(entry.getPhone(), 1);
            }
        }

        for(Entry entry : may) {
            if(mayPlayers.containsKey(entry.getPhone())) {
                int currCount = mayPlayers.get(entry.getPhone());
                mayPlayers.put(entry.getPhone(), ++currCount);
            } else {
                mayPlayers.put(entry.getPhone(), 1);
            }
        }

        frame.println("\n-> March player Phone : <Phone number> -> <occurrence>\n");
        marchPlayers.forEach((key, value) -> {
            frame.println("Phone : " + key + " -> " + value);
        });

        frame.println("\n-> April player Phone : <Phone number> -> <occurrence>\n");
        aprilPlayers.forEach((key, value) -> {
            frame.println("Phone : " + key + " -> " + value);
        });

        frame.println("\n-> May player Phone : <Phone number> ->  <occurrence>\n");
        mayPlayers.forEach((key, value) -> {
            frame.println("Phone : " + key + " -> " + value);
        });
    }

    private static void splitEntries() {
        march = new ArrayList<Entry>();
        april = new ArrayList<Entry>();
        may = new ArrayList<Entry>();

        for(Entry entry : entries) {
            addEntry(entry);
        }

        frame.println("-> 1th month\n");
        march.forEach(e->frame.println(e.toString()));

        frame.println("\n-> 2nd month\n");
        april.forEach(e->frame.println(e.toString()));

        frame.println("\n-> 3rd month\n");
        may.forEach(e->frame.println(e.toString()));
    }

    private static void addEntry(Entry entry) {
        if(monthMap.keySet().size() == 0) {
            monthMap.put(entry.getMonth(), "march");
        }

        if(!monthMap.containsKey(entry.getMonth())) {
            if(monthMap.keySet().size() == 1) {
                monthMap.put(entry.getMonth(), "april");
            } else if(monthMap.keySet().size() == 2) {
                monthMap.put(entry.getMonth(), "may");
            }
        }

        if(monthMap.get(entry.getMonth()).equals("march"))
            march.add(entry);
        if(monthMap.get(entry.getMonth()).equals("april"))
            april.add(entry);
        if(monthMap.get(entry.getMonth()).equals("may"))
            may.add(entry);
    }

    private static void fillEntries(String sampleXlsxFilePath) throws IOException {

        FileInputStream inputStream = new FileInputStream(new File(sampleXlsxFilePath));
        entries = new ArrayList<Entry>();

        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet firstSheet = workbook.getSheetAt(0);
        Iterator<Row> iterator = firstSheet.iterator();

        boolean header = true;
        while (iterator.hasNext()) {

            Row nextRow = iterator.next();
            Iterator<Cell> cellIterator = nextRow.cellIterator();

            if(!header) {
                int colIndex = 0;
                Entry entry = new Entry();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();

                    switch (colIndex) {
                        case 0:
                            entry.setId(cell.getStringCellValue());
                            break;
                        case 1:
                            entry.setMonth(cell.getStringCellValue());
                            break;
                        case 2:
                            entry.setPhone(cell.getStringCellValue());
                            break;
                        case 3:
                            entry.setMail(cell.getStringCellValue());
                            break;
                        case 4:
                            entry.setAP(cell.getStringCellValue());
                            break;
                        case 5:
                            entry.setBN(cell.getStringCellValue());
                            break;
                        case 6:
                            entry.setBI(cell.getStringCellValue());
                            break;
                    }
                    colIndex++;
                }

                if(entry.getPhone() != null) {
                    frame.println(entry.toString());
                    entries.add(entry);
                }
            }

            header = false;


        }

        workbook.close();
        inputStream.close();
    }

    private static void fillMonths() {
        months = new ArrayList<String>();
        monthMap = new HashMap<>();

        months.add("Január,Januar,January");
        months.add("Február,Februar,February");
        months.add("Március,Marcius,March");
        months.add("Április,Aprilis,April");
        months.add("Május,Majus,May");
        months.add("Június,Junius,Juni");
        months.add("Július,Julius,July");
        months.add("Augusztus,Augusztus,August");
        months.add("Szeptember,Szeptember,September");
        months.add("Október,Oktober,October");
        months.add("November,November,November");
        months.add("December,December,December");
    }
}
