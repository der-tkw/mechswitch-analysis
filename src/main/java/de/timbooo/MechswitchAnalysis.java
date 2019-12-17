package de.timbooo;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class MechswitchAnalysis {
    private static Map<String, Switch> ownedSwitches = new HashMap<>();
    private static Map<String, Switch> missingSwitches = new HashMap<>();
    private static List<Switch> newSwitches = new ArrayList<>();
    private static Workbook wb;
    private static CreationHelper createHelper;
    private static Sheet boards;
    private static Sheet switches;
    private static File input;
    private static CellStyle cellStyleWithBorder;
    private static CellStyle cellStyleWithoutBorder;

    public static void main(String[] args) throws Exception {
        prepare(args);
        createCellStyles();

        collect();
        linkToBoard();
        linkToSwitches();
        stats();

        save();
    }

    private static void prepare(String[] args) throws IOException, InvalidFormatException {
        String path;
        if (args != null && args.length == 1) {
            path = args[0];
        } else {
            path = "C:\\Users\\Tim-PC\\Desktop\\Temp\\switches.xlsx";
        }

        input = new File(path);
        if (input == null || !input.isFile()) {
            throw new IllegalArgumentException(path + " is not a file");
        }

        wb = new XSSFWorkbook(input);
        createHelper = wb.getCreationHelper();

        boards = wb.getSheet("Boards");
        switches = wb.getSheet("Switches");

    }

    private static void createCellStyles() {
        cellStyleWithBorder = wb.createCellStyle();
        Font font = wb.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 10);
        font.setUnderline(XSSFFont.U_SINGLE);
        cellStyleWithBorder.setFont(font);
        cellStyleWithBorder.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyleWithBorder.setAlignment(HorizontalAlignment.CENTER);
        cellStyleWithBorder.setWrapText(true);

        cellStyleWithoutBorder = wb.createCellStyle();
        cellStyleWithoutBorder.cloneStyleFrom(cellStyleWithBorder);

        cellStyleWithBorder.setBorderBottom(BorderStyle.THIN);
        cellStyleWithBorder.setBorderTop(BorderStyle.THIN);
        cellStyleWithBorder.setBorderRight(BorderStyle.THIN);
        cellStyleWithBorder.setBorderLeft(BorderStyle.THIN);
    }

    private static void collect() {
        for (Row row : switches) {
            if (row.getRowNum() == 0) {
                // skip first row (header)
                continue;
            }
            if (row.getCell(1) == null) {
                // stop at the end
                break;
            }

            String name = row.getCell(3).getStringCellValue();
            if (row.getCell(0) == null || StringUtils.isEmpty(row.getCell(0).getStringCellValue())) {
                missingSwitches.put(name, new Switch(name, row.getRowNum()));
            } else {
                ownedSwitches.put(name, new Switch(name, row.getRowNum()));
            }
        }
    }

    private static void linkToBoard() {
        for (Row row : boards) {
            for (Cell cell : row) {
                String name = cell.getStringCellValue();
                if (cell.getCellType().equals(CellType.FORMULA) && !ownedSwitches.containsKey(name)) {
                    Switch s = missingSwitches.get(name);
                    Cell linkCell = switches.getRow(s.getSwitchRow()).createCell(0);
                    linkCell.setCellValue("Board Link");
                    createLink(linkCell, "'Boards'!" + cell.getAddress());
                    linkCell.setCellStyle(cellStyleWithoutBorder);

                    missingSwitches.remove(name);
                    ownedSwitches.put(name, s);
                    newSwitches.add(s);
                }
            }
        }
    }

    private static void linkToSwitches() {
        for (Row row : boards) {
            for (Cell cell : row) {
                String name = cell.getStringCellValue();
                if (cell.getCellType().equals(CellType.FORMULA)) {
                    createLink(cell, "'Switches'!" + switches.getRow(getSwitchRow(name)).getCell(3).getAddress());
                    cell.setCellStyle(cellStyleWithBorder);
                }
            }
        }
    }

    private static void stats() {
        System.out.println("\n####################");
        System.out.println("# OWNED SWITCHES   #");
        System.out.println("####################");
        ownedSwitches.keySet().stream().sorted().forEach(o -> System.out.println(o));

        if (!missingSwitches.isEmpty()) {
            System.out.println("\n####################");
            System.out.println("# MISSING SWITCHES #");
            System.out.println("####################");
            missingSwitches.keySet().stream().sorted().forEach(o -> System.out.println(o));
        }

        if (!newSwitches.isEmpty()) {
            System.out.println("\n####################");
            System.out.println("# NEW SWITCHES     #");
            System.out.println("####################");
            newSwitches.forEach(o -> System.out.println(o));
        }

        System.out.println("\n####################");
        System.out.println("# STATS            #");
        System.out.println("####################");
        System.out.println("Owned Switches: " + ownedSwitches.size());
        System.out.println("Missing Switches: " + missingSwitches.size());
    }

    private static void save() throws IOException {
        try (FileOutputStream fos = new FileOutputStream(new File(input.getParentFile(),
                FilenameUtils.removeExtension(input.getName()) + "_new.xlsx"))) {
            wb.write(fos);
        }
    }

    // #### HELPERS ####

    private static void createLink(Cell cell, String address) {
        Hyperlink link = createHelper.createHyperlink(HyperlinkType.DOCUMENT);
        link.setAddress(address);
        cell.setHyperlink(link);
    }

    private static int getSwitchRow(String name) {
        if (ownedSwitches.containsKey(name)) {
            return ownedSwitches.get(name).getSwitchRow();
        } else if (missingSwitches.containsKey(name)) {
            return missingSwitches.get(name).getSwitchRow();
        }
        throw new IllegalStateException("Unknown switch on board: " + name);
    }
}
