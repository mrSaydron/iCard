package ru.mrak.iCard;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static org.apache.poi.ss.usermodel.BorderStyle.DOTTED;
import static org.apache.poi.ss.usermodel.BorderStyle.NONE;
import static org.apache.poi.ss.usermodel.BorderStyle.THIN;
import static ru.mrak.iCard.Align.CENTER;
import static ru.mrak.iCard.Align.LEFT;
import static ru.mrak.iCard.Align.ROTATION;
import static ru.mrak.iCard.util.PixelUtil.pixel2WidthUnits;

/**
 * Создает уинформационно-удостоверяющий лист из переданных данных
 * и сохраняет его в формате Excel
 */
public class WorkBook {
    private Workbook wb;
    private Sheet sheet;
    private List<EntryDoc> entryDocs;
    private FProperties properties;

    //Константы для формирования листов
    private static final int FIRST_PAGE_HEIGHT = 36;
    private static final int FOLLOW_PAGE_HEIGHT = 34;
    private static final int FIRST_PAGE_CAPACITY = 6;
    private static final int FOLLOW_PAGE_CAPACITY = 7;

    //Шрифты используемые в оформлении
    private final Font ARIAL_8;
    private final Font ARIAL_9;
    private final Font ARIAL_10;
    private final Font ARIAL_11;
    private final Font ARIAL_16;
    private final Font ARIAL_20;

    /**
     * Создает книгу и лист, настраивает шрифты
     * @param sheetName - наименование листа
     */
    private WorkBook(String sheetName){
        //Создаю книгу и лист
        wb = new HSSFWorkbook();
        sheet = wb.createSheet(WorkbookUtil.createSafeSheetName(sheetName));

        //Создаю шрифты
        ARIAL_8 = wb.createFont();
        ARIAL_8.setFontHeightInPoints((short)8);
        ARIAL_8.setItalic(true);

        ARIAL_9 = wb.createFont();
        ARIAL_9.setFontHeightInPoints((short)9);
        ARIAL_9.setItalic(true);

        ARIAL_10 = wb.createFont();
        ARIAL_10.setFontHeightInPoints((short)10);
        ARIAL_10.setItalic(true);

        ARIAL_11 = wb.createFont();
        ARIAL_11.setFontHeightInPoints((short)11);
        ARIAL_11.setItalic(true);

        ARIAL_16 = wb.createFont();
        ARIAL_16.setFontHeightInPoints((short)16);
        ARIAL_16.setItalic(true);

        ARIAL_20 = wb.createFont();
        ARIAL_20.setFontHeightInPoints((short)20);
        ARIAL_20.setItalic(true);
    }

    /**
     * Заполняет книгу переданной информацией и сохраняет в файл
     * @param sheetName - наименвоание листа
     * @param path - путь по которому необходимо сохранить книгу
     * @param entryDocs - параметры хокументав для занесения в книгу
     * @param properties - конфигурирование записей в книгу
     */
    public static void writeBook(String sheetName, String path, List<EntryDoc> entryDocs, FProperties properties) {
        WorkBook workBook = new WorkBook(sheetName);
        workBook.entryDocs = entryDocs;
        workBook.properties = properties;
        System.out.println("Создание информационно-удостоверяющего листа");
        workBook.createWorkBook();
        System.out.println("Запись информационно-удостоверяющего листа");
        workBook.saveWorkBook(path);
    }

    /**
     * Сохраняет книгу на диск
     * @param path - путь по которому необходимо сохранить книгу
     */
    private void saveWorkBook(String path) {
        try(FileOutputStream fileOut = new FileOutputStream(path)) {
            wb.write(fileOut);
            fileOut.close();
        } catch (IOException e) {
            System.out.println("Не удалось записать информационно-удостоверяющий лист");
            System.exit(0);
        }

    }

    /**
     * Форматирует лист и заполняет его необходимой информацией
     */
    private void createWorkBook() {
        //Определяю колличество страниц
        int numberOfPage;
        if(entryDocs.size() <= FIRST_PAGE_CAPACITY) {
            numberOfPage = 1;
        } else {
            numberOfPage = 1 + (int)Math.ceil((entryDocs.size() - FIRST_PAGE_CAPACITY) / (FOLLOW_PAGE_CAPACITY * 1.0));
        }
        //Настройка печати
        wb.setPrintArea(0, "A1:Q" + (numberOfPage == 1 ? 36 : (36 + (numberOfPage - 1) * 34)));
        PrintSetup printSetup = sheet.getPrintSetup();
        sheet.setFitToPage(true);
        printSetup.setScale((short)93);
        printSetup.setFooterMargin(0);
        printSetup.setHeaderMargin(0);
        printSetup.setPaperSize(printSetup.A4_PAPERSIZE);
        sheet.setMargin(Sheet.BottomMargin, 0);
        sheet.setMargin(Sheet.TopMargin, 0);
        sheet.setMargin(Sheet.LeftMargin, 0);
        sheet.setMargin(Sheet.RightMargin, 0);
        sheet.setHorizontallyCenter(true);
        sheet.setVerticallyCenter(true);
        //Стили по умолчанию
        
        //Создаю листы
        createFirstPage();
        for (int i = 2; i <=  numberOfPage; i++) {
            createFollowPage(i);
            sheet.setRowBreak(35 + 34 * (i - 2));
        }
        //Записываю колличество страниц
        printPageNumber(numberOfPage);
        //Записываю обозначения документов
        for (int i = 1; i <= entryDocs.size(); i++) {
            printEntry(i);
        }
    }

    /**
     * Формирует на листе одно поле для записи параметров документа
     * @param entryNumber - номер записи
     */
    private void printEntry(int entryNumber) {
        int firstRowIndex;
        if(entryNumber <= FIRST_PAGE_CAPACITY) {
            firstRowIndex = 2 + 4 * (entryNumber - 1);
        } else {
            int pageNumber = 1 + (int)Math.ceil((entryNumber - FIRST_PAGE_CAPACITY) / (FOLLOW_PAGE_CAPACITY * 1.0));
            firstRowIndex = 14 + 6 * (pageNumber - 2) + 4 * (entryNumber - 1);
        }
        writeText("" + entryNumber, shiftAddress("C0", firstRowIndex), CENTER, ARIAL_10);
        writeText(entryDocs.get(entryNumber - 1).getDesignation(), shiftAddress("F0", firstRowIndex), CENTER, ARIAL_10);
        writeText(entryDocs.get(entryNumber - 1).getName(), shiftAddress("J0", firstRowIndex), CENTER, ARIAL_10);
        writeText(entryDocs.get(entryNumber - 1).getVersion(), shiftAddress("L0", firstRowIndex), CENTER, ARIAL_10);
        writeText(entryDocs.get(entryNumber - 1).getReleaseNumber(), shiftAddress("O0", firstRowIndex), CENTER, ARIAL_10);

        writeText("MD5", shiftAddress("C1", firstRowIndex), CENTER, ARIAL_10);
        writeText(entryDocs.get(entryNumber - 1).getMd5(), shiftAddress("F1", firstRowIndex), LEFT, ARIAL_10);

        writeText(entryDocs.get(entryNumber - 1).getFileName(), shiftAddress("C2", firstRowIndex), LEFT, ARIAL_10);
        if(entryDocs.get(entryNumber - 1).getAuthor() == null) {
            writeFormula("F31", shiftAddress("L2", firstRowIndex), CENTER, ARIAL_10);
        } else {
            writeText(entryDocs.get(entryNumber - 1).getAuthor(), shiftAddress("L2", firstRowIndex), CENTER, ARIAL_10);
        }

        writeText(entryDocs.get(entryNumber - 1).getSize(), shiftAddress("C3", firstRowIndex), CENTER, ARIAL_10);
        writeText(entryDocs.get(entryNumber - 1).getDate(), shiftAddress("J3", firstRowIndex), CENTER, ARIAL_10);
    }

    /**
     * Заполняет колличество листов форматке
     * @param numberOfPage - общее количество листов
     */
    private void printPageNumber(int numberOfPage) {
        if(numberOfPage == 1) {
            writeText("1","P32", CENTER, ARIAL_9);
        } else {
            writeText("1","N32", CENTER, ARIAL_9);
            writeText("" + numberOfPage,"P32", CENTER, ARIAL_9);
        }
    }

    /**
     * Формирует первыл лист УЛ
     */
    private void createFirstPage() {
        //Настраиваю ширину столбцов и строк
        sheet.setColumnWidth(0, pixel2WidthUnits(19));//A
        sheet.setColumnWidth(1, pixel2WidthUnits(26));//B
        sheet.setColumnWidth(2, pixel2WidthUnits(26));//C
        sheet.setColumnWidth(3, pixel2WidthUnits(29));//D
        sheet.setColumnWidth(4, pixel2WidthUnits(9));//E
        sheet.setColumnWidth(5, pixel2WidthUnits(86));//F
        sheet.setColumnWidth(6, pixel2WidthUnits(57));//G
        sheet.setColumnWidth(7, pixel2WidthUnits(38));//H
        sheet.setColumnWidth(8, pixel2WidthUnits(19));//I
        sheet.setColumnWidth(9, pixel2WidthUnits(246));//J
        sheet.setColumnWidth(10, pixel2WidthUnits(19));//K
        sheet.setColumnWidth(11, pixel2WidthUnits(19));//L
        sheet.setColumnWidth(12, pixel2WidthUnits(19));//M
        sheet.setColumnWidth(13, pixel2WidthUnits(36));//N
        sheet.setColumnWidth(14, pixel2WidthUnits(20));//O
        sheet.setColumnWidth(15, pixel2WidthUnits(36));//P
        sheet.setColumnWidth(16, pixel2WidthUnits(40));//Q

        //Настраиваю высоту столбцов
        sheet.createRow(0).setHeightInPoints(45f);
        for(int i = 1; i < 27; i++)
            sheet.createRow(i).setHeightInPoints(25.5f);
        for(int i = 27; i < 36; i++)
            sheet.createRow(i).setHeightInPoints(15f);

        //Объединяю ячейки
        //Форматка
        //Первичное применение
        regionMerge("A1:A7", THIN);
        regionMerge("B1:B7", THIN);
        //Справочный No
        regionMerge("A8:A14", THIN);
        regionMerge("B8:B14", THIN);
        //Подпись и дата
        regionMerge("A16:A19", THIN);
        regionMerge("B16:B19", THIN);
        //Инв. No дубл.
        regionMerge("A20:A22", THIN);
        regionMerge("B20:B22", THIN);
        //Взаим. инв No.
        regionMerge("A23:A25", THIN);
        regionMerge("B23:B25", THIN);
        //Подп. и дата
        regionMerge("A26:A30", THIN);
        regionMerge("B26:B30", THIN);
        //Инв. No подп.
        regionMerge("A31:A35", THIN);
        regionMerge("B31:B35", THIN);
        //Извещения
        regionMerge("C28", DOTTED, THIN, THIN, THIN);
        regionMerge("D28:E28", DOTTED, THIN, THIN, THIN);
        regionMerge("F28", DOTTED, THIN, THIN, THIN);
        regionMerge("G28", DOTTED, THIN, THIN, THIN);
        regionMerge("H28", DOTTED, THIN, THIN, THIN);
        regionMerge("C29", THIN, DOTTED, THIN, THIN);
        regionMerge("D29:E29", THIN, DOTTED, THIN, THIN);
        regionMerge("F29", THIN, DOTTED, THIN, THIN);
        regionMerge("G29", THIN, DOTTED, THIN, THIN);
        regionMerge("H29", THIN, DOTTED, THIN, THIN);
        //Подписи
        regionMerge("C30", THIN);
        regionMerge("D30:E30", THIN);
        regionMerge("F30", THIN);
        regionMerge("G30", THIN);
        regionMerge("H30", THIN);
        regionMerge("C31:E31", DOTTED, THIN, THIN, THIN);
        regionMerge("F31", DOTTED, THIN, THIN, THIN);
        regionMerge("G31", DOTTED, THIN, THIN, THIN);
        regionMerge("H31", DOTTED, THIN, THIN, THIN);
        for (int i = 32; i < 35; i++) {
            regionMerge("C" + i + ":E" + i, DOTTED, DOTTED, THIN, THIN);
            regionMerge("F" + i, DOTTED, DOTTED, THIN, THIN);
            regionMerge("G" + i, DOTTED, DOTTED, THIN, THIN);
            regionMerge("H" + i, DOTTED, DOTTED, THIN, THIN);
        }
        regionMerge("C35:E35", THIN, DOTTED, THIN, THIN);
        regionMerge("F35", THIN, DOTTED, THIN, THIN);
        regionMerge("G35", THIN, DOTTED, THIN, THIN);
        regionMerge("H35", THIN, DOTTED, THIN, THIN);
        //Обозначение
        regionMerge("I28:Q30", THIN);
        //Наименование
        regionMerge("I31:J35", THIN);
        //Литера
        regionMerge("K31:M31", THIN);
        regionMerge("K32", THIN);
        regionMerge("L32", THIN);
        regionMerge("M32", THIN);
        //Листы
        regionMerge("N31:O31", THIN);
        regionMerge("N32:O32", THIN);
        regionMerge("P31:Q31", THIN);
        regionMerge("P32:Q32", THIN);
        //Компания
        regionMerge("K33:Q35", THIN);

        //Заголовок
        regionMerge("C1:E1", THIN);
        regionMerge("F1:I1", THIN);
        regionMerge("J1:K1", THIN);
        regionMerge("L1:N1", THIN);
        regionMerge("O1:Q1", THIN);

        //Записи
        for (int i = 0; i < FIRST_PAGE_CAPACITY ; i++) {
            entry(2 + i * 4);
        }

        //Рамка
        regionMerge("Q26", NONE, THIN, NONE, THIN);
        regionMerge("Q27", THIN, NONE, NONE, THIN);

        //Надписи
        writeText("ПАО \"Морион\"", "K33", CENTER, ARIAL_16);
        writeText("Изм.", "C30", CENTER, ARIAL_9);
        writeText("Лист", "D30", CENTER, ARIAL_9);
        writeText("№ докум.", "F30", CENTER, ARIAL_9);
        writeText("Подп.", "G30", CENTER, ARIAL_9);
        writeText("Дата", "H30", CENTER, ARIAL_9);
        writeText("Лит.", "K31", CENTER, ARIAL_9);
        writeText("Лист", "N31", CENTER, ARIAL_9);
        writeText("Листов.", "P31", CENTER, ARIAL_9);
        writeText("Разраб.", "C31", LEFT, ARIAL_9);
        writeText("Пров.", "C32", LEFT, ARIAL_9);
        writeText("Н. контр.", "C34", LEFT, ARIAL_9);
        writeText("Утв.", "C35", LEFT, ARIAL_9);
        writeText("№ поз./\nHash", "C1", CENTER, ARIAL_9);
        writeText("Обозначение\nдокумента", "F1", CENTER, ARIAL_9);
        writeText("Наименование изделия\nнаименование документа", "J1", CENTER, ARIAL_9);
        writeText("Версия/\nизменение", "L1", CENTER, ARIAL_8);
        writeText("Номер\nрелиза", "O1", CENTER, ARIAL_9);
        writeText("Информационно-\nудостоверяющий\nлист", "I31", CENTER, ARIAL_16);
        writeText("Перв. примен.", "A1", ROTATION, ARIAL_9);
        writeText("Справ. №", "A8", ROTATION, ARIAL_9);
        writeText("Подп. и дата", "A16", ROTATION, ARIAL_9);
        writeText("Инв. № дубл.", "A20", ROTATION, ARIAL_9);
        writeText("Взам. инв №", "A23", ROTATION, ARIAL_9);
        writeText("Подп. и дата", "A26", ROTATION, ARIAL_9);
        writeText("Инв. № подл.", "A31", ROTATION, ARIAL_9);
        writeText("Копировал:", "I36", LEFT, ARIAL_9);
        writeText("Формат: А4", "K36", LEFT, ARIAL_9);
        writeText(properties.getAuthor(), "F31", LEFT, ARIAL_9);
        writeText(properties.getChecked(), "F32", LEFT, ARIAL_9);
        writeText(properties.getApproved(), "F35", LEFT, ARIAL_9);
        writeText("ИГУЛ.000000.000-УЛ", "I28", CENTER, ARIAL_20);

        //Устанавливаю формат ячеек
        writeText("", "F34", LEFT, ARIAL_9);
        writeText("", "K32", CENTER, ARIAL_9);
        writeText("", "L32", CENTER, ARIAL_9);
        writeText("", "M32", CENTER, ARIAL_9);
        writeText("", "B1", ROTATION, ARIAL_9);
    }

    /**
     * Формирует второй и последующий листы УЛ
     * @param pageNumber - номер листа
     */
    private void createFollowPage(int pageNumber) {
        if(pageNumber < 2) throw new RuntimeException ("Неверный номер для следующей страницы");
        int firstRowIndex = FIRST_PAGE_HEIGHT + (FOLLOW_PAGE_HEIGHT * (pageNumber - 2));

        //Настраиваю высоту столбцов
        sheet.createRow(firstRowIndex).setHeightInPoints(45f);
        for(int i = firstRowIndex + 1; i < firstRowIndex + 30; i++)
            sheet.createRow(i).setHeightInPoints(25.5f);
        for(int i = firstRowIndex + 30; i < firstRowIndex + 34; i++)
            sheet.createRow(i).setHeightInPoints(15f);

        //Объединяю ячейки
        //Форматка
        //Подпись и дата
        regionMerge(shiftRegion("A16:A19", firstRowIndex), THIN);
        regionMerge(shiftRegion("B16:B19", firstRowIndex), THIN);
        //Инв. No дубл.
        regionMerge(shiftRegion("A20:A22", firstRowIndex), THIN);
        regionMerge(shiftRegion("B20:B22", firstRowIndex), THIN);
        //Взаим. инв No.
        regionMerge(shiftRegion("A23:A25", firstRowIndex), THIN);
        regionMerge(shiftRegion("B23:B25", firstRowIndex), THIN);
        //Подп. и дата
        regionMerge(shiftRegion("A26:A29", firstRowIndex), THIN);
        regionMerge(shiftRegion("B26:B29", firstRowIndex), THIN);
        //Инв. No подп.
        regionMerge(shiftRegion("A30:A33", firstRowIndex), THIN);
        regionMerge(shiftRegion("B30:B33", firstRowIndex), THIN);
        //Извещения
        regionMerge(shiftRegion("C31", firstRowIndex), DOTTED, THIN, THIN, THIN);
        regionMerge(shiftRegion("D31:E31", firstRowIndex), DOTTED, THIN, THIN, THIN);
        regionMerge(shiftRegion("F31", firstRowIndex), DOTTED, THIN, THIN, THIN);
        regionMerge(shiftRegion("G31", firstRowIndex), DOTTED, THIN, THIN, THIN);
        regionMerge(shiftRegion("H31", firstRowIndex), DOTTED, THIN, THIN, THIN);
        regionMerge(shiftRegion("C32", firstRowIndex), THIN, DOTTED, THIN, THIN);
        regionMerge(shiftRegion("D32:E32", firstRowIndex), THIN, DOTTED, THIN, THIN);
        regionMerge(shiftRegion("F32", firstRowIndex), THIN, DOTTED, THIN, THIN);
        regionMerge(shiftRegion("G32", firstRowIndex), THIN, DOTTED, THIN, THIN);
        regionMerge(shiftRegion("H32", firstRowIndex), THIN, DOTTED, THIN, THIN);
        regionMerge(shiftRegion("C33", firstRowIndex), THIN);
        regionMerge(shiftRegion("D33:E33", firstRowIndex), THIN);
        regionMerge(shiftRegion("F33", firstRowIndex), THIN);
        regionMerge(shiftRegion("G33", firstRowIndex), THIN);
        regionMerge(shiftRegion("H33", firstRowIndex), THIN);
        //Обозначение
        regionMerge(shiftRegion("I31:P33", firstRowIndex), THIN);
        //Листы
        regionMerge(shiftRegion("Q31", firstRowIndex), THIN);
        regionMerge(shiftRegion("Q32:Q33", firstRowIndex), THIN);
        //Заголовок
        regionMerge(shiftRegion("C1:E1", firstRowIndex), THIN);
        regionMerge(shiftRegion("F1:I1", firstRowIndex), THIN);
        regionMerge(shiftRegion("J1:K1", firstRowIndex), THIN);
        regionMerge(shiftRegion("L1:N1", firstRowIndex), THIN);
        regionMerge(shiftRegion("O1:Q1", firstRowIndex), THIN);

        //Записи
        for (int i = 0; i < FOLLOW_PAGE_CAPACITY ; i++) {
            entry(2 + i * 4 + firstRowIndex);
        }

        //Рамка
        regionMerge(shiftRegion("Q30", firstRowIndex), THIN, THIN, NONE, THIN);

        //Надписи
        writeText("Изм.", shiftAddress("C33", firstRowIndex), CENTER, ARIAL_9);
        writeText("Лист", shiftAddress("D33", firstRowIndex), CENTER, ARIAL_9);
        writeText("№ докум.", shiftAddress("F33", firstRowIndex), CENTER, ARIAL_9);
        writeText("Подп.", shiftAddress("G33", firstRowIndex), CENTER, ARIAL_9);
        writeText("Дата", shiftAddress("H33", firstRowIndex), CENTER, ARIAL_9);
        writeText("№ поз./\nHash", shiftAddress("C1", firstRowIndex), CENTER, ARIAL_9);
        writeText("Обозначение\nдокумента", shiftAddress("F1", firstRowIndex), CENTER, ARIAL_9);
        writeText("Наименование изделия\nнаименование документа", shiftAddress("J1", firstRowIndex), CENTER, ARIAL_9);
        writeText("Версия/\nизменение", shiftAddress("L1", firstRowIndex), CENTER, ARIAL_8);
        writeText("Номер\nрелиза", shiftAddress("O1", firstRowIndex), CENTER, ARIAL_9);
        writeText("Подп. и дата", shiftAddress("A16", firstRowIndex), ROTATION, ARIAL_9);
        writeText("Инв. № дубл.", shiftAddress("A20", firstRowIndex), ROTATION, ARIAL_9);
        writeText("Взам. инв №", shiftAddress("A23", firstRowIndex), ROTATION, ARIAL_9);
        writeText("Подп. и дата", shiftAddress("A26", firstRowIndex), ROTATION, ARIAL_9);
        writeText("Инв. № подл.", shiftAddress("A30", firstRowIndex), ROTATION, ARIAL_9);
        writeText("Копировал:", shiftAddress("I34", firstRowIndex), LEFT, ARIAL_9);
        writeText("Формат: А4", shiftAddress("K34", firstRowIndex), LEFT, ARIAL_9);
        writeText("Лист", shiftAddress("Q31", firstRowIndex), CENTER, ARIAL_9);
        writeText("" + pageNumber, shiftAddress("Q32", firstRowIndex), CENTER, ARIAL_11);
        writeFormula("I28", shiftAddress("I31", firstRowIndex), CENTER, ARIAL_20);

        //Область печати
        //wb.setPrintArea(0, "A1:Q36");
    }

    /**
     * Ихменяет запись региона смещая его на заданное количество строк
     * @param region - запись региона
     * @param shift - количество строк на которое смещается регион
     * @return новую запист региона
     */
    private String shiftRegion(String region, int shift) {
        if(region.contains(":")) {
            String[] addresses = region.split(":");
            return shiftAddress(addresses[0], shift) + ":" + shiftAddress(addresses[1], shift);
        } else {
            return shiftAddress(region, shift);
        }
    }

    /**
     * хменяет запись адреса смещая его на заданное количество строк
     * @param address - запись адреса
     * @param shift - количество строк на которое смещается адрес
     * @return запись адреса
     */
    private String shiftAddress(String address, int shift) {
        Pattern pattern = Pattern.compile("\\d+");
        Matcher matcher = pattern.matcher(address);
        String row;
        if(matcher.find()) {
            row = matcher.group();
        }
        else
            throw new RuntimeException("Не правильно задан аддресс");
        pattern = Pattern.compile("\\D+");
        matcher = pattern.matcher(address);
        String cell;
        if(matcher.find()) {
            cell = matcher.group();
        }
        else
            throw new RuntimeException("Не правильно задан аддресс");
        //String ad_ = address.replaceAll("[A-Z]", "");
        int rowIndex = Integer.parseInt(row);
        return cell + (rowIndex + shift);
    }

    /**
     * Обрамяет регион указанной рамкой
     * @param region - регион
     * @param borderStyle - рамка
     */
    private void regionMerge(String region, BorderStyle borderStyle) {
        regionMerge(region, borderStyle, borderStyle, borderStyle, borderStyle);
    }

    /**
     * Обрамляет регион рамкой с указанием рамки и стороны
     * @param region - регион
     * @param borderBottom - рамка снизу
     * @param borderTop - рамка сверху
     * @param borderLeft - рамка слева
     * @param borderRight - рамка справа
     */
    private void regionMerge(String region, BorderStyle borderBottom, BorderStyle borderTop, BorderStyle borderLeft, BorderStyle borderRight) {
        if(region.contains(":")) {
            CellRangeAddress cellRangeAddress = CellRangeAddress.valueOf(region);
            sheet.addMergedRegion(cellRangeAddress);
            RegionUtil.setBorderBottom(borderBottom, cellRangeAddress, sheet);
            RegionUtil.setBorderTop(borderTop, cellRangeAddress, sheet);
            RegionUtil.setBorderLeft(borderLeft, cellRangeAddress, sheet);
            RegionUtil.setBorderRight(borderRight, cellRangeAddress, sheet);
        } else {
            CellAddress cellAddress = new CellAddress(region);
            Row row = sheet.getRow(cellAddress.getRow());
            Cell cell = row.getCell(cellAddress.getColumn());
            if(cell == null) cell = row.createCell(cellAddress.getColumn());
            CellStyle style = wb.createCellStyle();
            style.setBorderBottom(borderBottom);
            style.setBorderTop(borderTop);
            style.setBorderLeft(borderLeft);
            style.setBorderRight(borderRight);
            cell.setCellStyle(style);
        }
    }

    /**
     * Записывает параметры документа на страницу
     * @param rowIndex - первая строка в которую производится запись
     */
    private void entry(int rowIndex) {
        regionMerge("C" + rowIndex + ":E" + rowIndex, DOTTED, THIN, THIN, THIN);
        regionMerge("F" + rowIndex + ":I" + rowIndex, DOTTED, THIN, THIN, THIN);
        regionMerge("J" + rowIndex + ":K" + rowIndex, DOTTED, THIN, THIN, THIN);
        regionMerge("L" + rowIndex + ":N" + rowIndex, DOTTED, THIN, THIN, THIN);
        regionMerge("O" + rowIndex + ":Q" + rowIndex, DOTTED, THIN, THIN, THIN);
        rowIndex++;
        regionMerge("C" + rowIndex + ":E" + rowIndex, DOTTED, DOTTED, THIN, THIN);
        regionMerge("F" + rowIndex + ":K" + rowIndex, DOTTED, DOTTED, THIN, THIN);
        regionMerge("L" + rowIndex + ":Q" + rowIndex, DOTTED, DOTTED, THIN, THIN);
        rowIndex++;
        regionMerge("C" + rowIndex + ":K" + rowIndex, DOTTED, DOTTED, THIN, THIN);
        regionMerge("L" + rowIndex + ":Q" + rowIndex, DOTTED, DOTTED, THIN, THIN);
        rowIndex++;
        regionMerge("C" + rowIndex + ":I" + rowIndex, THIN, DOTTED, THIN, THIN);
        regionMerge("J" + rowIndex + ":K" + rowIndex, THIN, DOTTED, THIN, THIN);
        regionMerge("L" + rowIndex + ":Q" + rowIndex, THIN, DOTTED, THIN, THIN);
    }

    /**
     * Пишет текст в указанную ячейку
     * @param text - текст
     * @param cellAddress - адрес ячейки
     * @param align - выравнивание
     * @param font - шрифт
     */
    private void writeText(String text, String cellAddress, Align align, Font font) {
        Cell cell = setCellStyle(cellAddress, align, font, text.contains("\n"));
        cell.setCellValue(text);
    }

    /**
     * Пишет формулу в указанную ячейку
     * @param formula - формулы
     * @param cellAddress - адрес ячейки
     * @param align - выравнивание
     * @param font - шрифт
     */
    private void writeFormula(String formula, String cellAddress, Align align, Font font) {
        Cell cell = setCellStyle(cellAddress, align, font, false);
        cell.setCellFormula(formula);
    }

    /**
     * Устанавливает стиль для ячейки
     * @param cellAddress - адрес ячейки
     * @param align -выравнивание
     * @param font - шрифт
     * @param wrap - однострочный/многострочный
     * @return ячейку
     */
    private Cell setCellStyle(String cellAddress, Align align, Font font, boolean wrap) {
        Cell cell = getCell(cellAddress);
        Map<String, Object> properties = new HashMap<String, Object>();
        switch (align) {
            case LEFT:
                properties.put(CellUtil.ALIGNMENT, HorizontalAlignment.LEFT);
                break;
            case CENTER:
                properties.put(CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);
                break;
            case ROTATION:
                properties.put(CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);
                properties.put(CellUtil.ROTATION, new Short((short)90));
                break;
        }
        properties.put(CellUtil.VERTICAL_ALIGNMENT, VerticalAlignment.CENTER);
        if(wrap) properties.put(CellUtil.WRAP_TEXT, true);
        properties.put(CellUtil.FONT, font.getIndex());
        CellUtil.setCellStyleProperties(cell, properties);
        return cell;
    }

    /**
     * Возврящает ячейку по указанному адресу
     * @param cellAddress - адрес
     * @return ячейка
     */
    private Cell getCell(String cellAddress) {
        CellAddress address = new CellAddress(cellAddress);
        Row row = sheet.getRow(address.getRow());
        Cell cell = row.getCell(address.getColumn());
        if(cell == null) cell = row.createCell(address.getColumn());
        return cell;
    }
}

/**
 * Выравнивания
 */
enum Align {
    LEFT,
    CENTER,
    ROTATION
}