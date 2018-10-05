package com.kazurayam.ksworkbook

import java.nio.file.Files
import java.nio.file.Path

import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.CreationHelper
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel.RichTextString
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook

/**
 * KSWorkbook wraps an instance of Apache POI Workbook (Excel spreadsheet).
 * KSWorkbook is designed to be used inside a Katalon Studio project.
 *
 */
class KSWorkbook {

    private Path excel_
    private Workbook wb_

    /**
     * constructor
     *
     * @param excel
     */
    KSWorkbook(Path excel) {
        excel_ = excel
        wb_ = findWorkbook(excel, DoIfNotPresent.CREATE)
    }

    /**
     *
     */
    void close() {
        if (excel_ != null) {
            if (wb_ != null) {
                serialize(wb_, excel_)
            } else {
                throw new IllegalStateException("workbook_ is null")
            }
        } else {
            throw new IllegalStateException("excel_ is null")
        }
    }

    /**
     *
     * @return
     */
    List<String> getSheetNames() {
        List<String> sheetNames = new ArrayList<String>()
        for (int i = 0; i < wb_.getNumberOfSheets(); i++) {
            sheetNames.add(wb_.getSheetName(i))
        }
        return sheetNames
    }

    /**
     *
     * @return 0 if there is 1 row contained. 9 if there are 10 rows contained
     */
    int getLastRowNum(String sheetName,
            DoIfNotPresent flowControl = DoIfNotPresent.STOP) {
        Sheet sheet = findSheet(sheetName, flowControl)
        return sheet.getLastRowNum()
    }

    /**
     * @return 0 if there is 1 cell contained. 9 if there are 10 cells contained
     */
    int getLastCellNum(String sheetName, int rownum,
            DoIfNotPresent flowControl = DoIfNotPresent.STOP) {
        Row row = findRow(sheetName, rownum, flowControl)
        return row.getLastCellNum()
    }


    /**
     *
     * @param sheetName
     * @param rownum
     * @param cellnum
     * @return
     */
    String readCell(String sheetName, int rownum, int cellnum,
            DoIfNotPresent flowControl = DoIfNotPresent.IGNORE) {
        Cell cell = findCell(sheetName, rownum, cellnum, flowControl)
        if (cell != null) {
            String result
            switch (getCellType(sheetName, rownum, cellnum)) {
                case Cell.CELL_TYPE_NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        result = cell.getDateCellValue()
                    } else {
                        result = cell.getNumericCellValue()
                    }
                    break
                case Cell.CELL_TYPE_STRING:
                    result = cell.getStringCellValue()
                    break
                case Cell.CELL_TYPE_FORMULA:
                    result = cell.getCellFormula()
                    break
                case Cell.CELL_TYPE_BLANK:
                    result = ''
                    break
                case Cell.CELL_TYPE_BOOLEAN:
                    result = cell.getBooleanCellValue()
                    break
                case Cell.CELL_TYPE_ERROR:
                    result = cell.getErrorCellValue()
                    break
                default:
                    break
            }
            return result
        } else {
            throw new KSWorkbookException("#getStringCellValue(\'${sheetName}\', rownum, cellnum) is not found")
        }
    }

    void writeCell(String sheetName, int rownum, int cellnum, boolean value,
            DoIfNotPresent flowControl = DoIfNotPresent.CREATE) {
        Cell cell = findCell(sheetName, rownum, cellnum, flowControl)
        cell.setCellValue(value)
    }
    void writeCell(String sheetName, int rownum, int cellnum, Calendar value,
            DoIfNotPresent flowControl = DoIfNotPresent.CREATE) {
        Cell cell = findCell(sheetName, rownum, cellnum, flowControl)
        cell.setCellValue(value)
        cell.setCellStyle(createDateCellStyle(wb_))
    }
    void writeCell(String sheetName, int rownum, int cellnum, Date value,
            DoIfNotPresent flowControl = DoIfNotPresent.CREATE) {
        Cell cell = findCell(sheetName, rownum, cellnum, flowControl)
        cell.setCellValue(value)
        cell.setCellStyle(createDateCellStyle(wb_))
    }
    void writeCell(String sheetName, int rownum, int cellnum, double value,
            DoIfNotPresent flowControl = DoIfNotPresent.CREATE) {
        Cell cell = findCell(sheetName, rownum, cellnum, flowControl)
        cell.setCellValue(value)
    }
    void writeCell(String sheetName, int rownum, int cellnum, RichTextString value,
            DoIfNotPresent flowControl = DoIfNotPresent.CREATE) {
        Cell cell = findCell(sheetName, rownum, cellnum, flowControl)
        cell.setCellValue(value)
    }
    void writeCell(String sheetName, int rownum, int cellnum, String value,
            DoIfNotPresent flowControl = DoIfNotPresent.CREATE) {
        Cell cell = findCell(sheetName, rownum, cellnum, flowControl)
        cell.setCellValue(value)
    }






    /** ------------------ private methos ---------------------------------- */

    /**
     * @return
     * 0 : Cell.CELL_TYPE_NUMERIC
     * 1 : Cell.CELL_TYPE_STRING
     * 2 : Cell.CELL_TYPE_FORMULA
     * 3 : Cell.CELL_TYPE_BLANK
     * 4 : Cell.CELL_TYPE_BOOLEAN
     * 5 : Cell.CELL_TYPE_ERROR
     */
    private int getCellType(String sheetName, int rownum, int cellnum) {
        Cell cell = findCell(sheetName, rownum, cellnum, DoIfNotPresent.STOP)
        return cell.getCellType()
    }

    private Cell findCell(String sheetName, int rownum, int cellnum,
            DoIfNotPresent flowControl = DoIfNotPresent.CREATE) {
        Sheet sheet = findSheet(sheetName, flowControl)
        Row row = findRow(sheetName, rownum, flowControl)
        Cell cell = row.getCell(cellnum)
        if (cell == null) {
            if (flowControl == DoIfNotPresent.CREATE) {
                row.createCell(cellnum)
                cell = row.getCell(cellnum)
            } else if (flowControl == DoIfNotPresent.STOP) {
                throw new KSWorkbookException("Sheet \'${sheet}\' Row(${rownum}) Cell(${cellnum}) is not present")
            }
        }
        return cell
    }

    private Row findRow(String sheetName, int rownum,
            DoIfNotPresent flowControl = DoIfNotPresent.CREATE) {
        Sheet sheet = findSheet(sheetName, flowControl)
        Row row = sheet.getRow(rownum)
        if (row == null) {
            if (flowControl == DoIfNotPresent.CREATE) {
                row = sheet.createRow(rownum)
            } else if (flowControl == DoIfNotPresent.STOP) {
                throw new KSWorkbookException("Sheet \'${sheet}\' Row(${rownum}) is not present")
            }
        }
        return row
    }

    /**
     *
     */
    private Sheet findSheet(String sheetName,
            DoIfNotPresent flowControl = DoIfNotPresent.CREATE) {
        Sheet sheet = wb_.getSheet(sheetName)
        if (sheet == null) {
            if (flowControl == DoIfNotPresent.CREATE) {
                wb_.createSheet(sheetName)
                sheet = wb_.getSheet(sheetName)
            } else if (flowControl == DoIfNotPresent.STOP) {
                throw new KSWorkbookException("Sheet \'${sheetName}\' is not present")
            }
        }
        return sheet
    }

    /**
     *
     */
    private static Workbook findWorkbook(Path excel,
            DoIfNotPresent flowControl = DoIfNotPresent.CREATE) {
        HSSFWorkbook wb
        if (Files.exists(excel)) {
            FileInputStream fis = new FileInputStream(excel.toFile())
            wb = new HSSFWorkbook(fis)
        } else {
            if (flowControl == DoIfNotPresent.CREATE) {
                wb = new HSSFWorkbook()
            } else if (flowControl == DoIfNotPresent.STOP) {
                throw new KSWorkbookException("${excel.toString()} is not present")
            }
        }
        return (Workbook)wb
    }

    /**
     *
     */
    private static CellStyle createDateCellStyle(Workbook wb) {
        CreationHelper createHelper = wb.getCreationHelper()
        CellStyle cellStyle = wb.createCellStyle()
        short style = createHelper.createDataFormat().getFormat("yyyy/mm/dd h:mm")
        cellStyle.setDataFormat(style)
        return cellStyle
    }

    /**
     *
     */
    private static void serialize(Workbook workbook, Path excel) {
        if (!Files.exists(excel.getParent())) {
            Files.createDirectories(excel.getParent())
        }
        FileOutputStream fos = new FileOutputStream(excel.toFile())
        workbook.write(fos)
        fos.flush()
        fos.close()
    }

    /**
     *
     */
    enum DoIfNotPresent {
        CREATE,
        STOP,
        IGNORE
    }
}