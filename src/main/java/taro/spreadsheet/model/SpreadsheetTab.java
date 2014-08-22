package taro.spreadsheet.model;

import static com.google.common.collect.Maps.newHashMap;
import static taro.spreadsheet.model.SpreadsheetCellStyle.DEFAULT;

import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

public class SpreadsheetTab {

  private final SpreadsheetWorkbook workbook;
  private final Sheet sheet;
  private final Map<String, SpreadsheetCell> cells = newHashMap();
  private Drawing drawing;

  private int highestModifiedCol = -1;
  private int highestModifiedRow = -1;

  public SpreadsheetTab(final SpreadsheetWorkbook workbook, final Sheet sheet) {
    this.workbook = workbook;
    this.sheet = sheet;
  }

  public SpreadsheetTab(final SpreadsheetWorkbook workbook, final String title) {
    this.workbook = workbook;
    this.sheet = workbook.getPoiWorkbook().createSheet(title);
  }

  public static String getCellAddress(final int row, final int col) {
    return CellReference.convertNumToColString(col) + (row + 1);
  }

  public void addPicture(final int row, final int col, final byte[] bytes, final int pictureType) {
    if (drawing == null) {
      drawing = sheet.createDrawingPatriarch();
    }

    final int pictureIndex = workbook.getPoiWorkbook().addPicture(bytes, pictureType);
    // add a picture shape
    final ClientAnchor anchor = workbook.getPoiWorkbook().getCreationHelper().createClientAnchor();
    // set top-left corner of the picture,
    // subsequent call of Picture#resize() will operate relative to it
    anchor.setCol1(col);
    anchor.setRow1(row);

    final Picture pict = drawing.createPicture(anchor, pictureIndex);
    // auto-size picture relative to its top-left corner
    pict.resize();
  }

  public void addPicture(final String cellAddress, final byte[] bytes, final int pictureType) {
    final CellReference cellRef = new CellReference(cellAddress);
    addPicture(cellRef.getRow(), cellRef.getCol(), bytes, pictureType);
  }

  public void addSpacer() {
    sheet.setColumnWidth(0, 768);
  }

  public void autosizeCols() {
    for (int col = 0; col <= highestModifiedCol; col++) {
      sheet.autoSizeColumn(col, true);
    }
  }

  public void autoSizeRow(final int row) {
    float tallestCell = -1;
    for (int col = 0; col <= highestModifiedCol; col++) {
      final SpreadsheetCell cell = getCell(row, col);
      final int fontSize = cell.getFontSizeInPoints();
      final Cell poiCell = cell.getPoiCell();
      if (poiCell.getCellType() == Cell.CELL_TYPE_STRING) {
        final String value = poiCell.getStringCellValue();
        int numLines = 1;
        for (int i = 0; i < value.length(); i++) {
          if (value.charAt(i) == '\n')
            numLines++;
        }
        final float cellHeight = computeRowHeightInPoints(fontSize, numLines);
        if (cellHeight > tallestCell) {
          tallestCell = cellHeight;
        }
      }
    }

    final float defaultRowHeightInPoints = sheet.getDefaultRowHeightInPoints();
    float rowHeight = tallestCell;
    if (rowHeight < defaultRowHeightInPoints + 1) {
      rowHeight = -1; // resets to the default
    }

    sheet.getRow(row).setHeightInPoints(rowHeight);
  }

  public void autosizeRows() {
    for (int row = 0; row <= highestModifiedRow; row++) {
      autoSizeRow(row);
    }
  }

  public void autosizeRowsAndCols() {
    autosizeCols();
    autosizeRows();
  }

  public float computeRowHeightInPoints(final int fontSizeInPoints, final int numLines) {
    // a crude approximation of what excel does
    final float defaultRowHeightInPoints = sheet.getDefaultRowHeightInPoints();
    float lineHeightInPoints = 1.3f * fontSizeInPoints;
    if (lineHeightInPoints < defaultRowHeightInPoints + 1) {
      lineHeightInPoints = defaultRowHeightInPoints;
    }
    float rowHeightInPoints = lineHeightInPoints * numLines;
    rowHeightInPoints = Math.round(rowHeightInPoints * 4) / 4f; // round to the nearest 0.25
    return rowHeightInPoints;
  }

  public void forceAutosizeRows() {
    highestModifiedRow = sheet.getLastRowNum();
    for (int row = 0; row <= highestModifiedRow; row++) {
      final short rowLastCellNum = sheet.getRow(row).getLastCellNum();
      if (rowLastCellNum > highestModifiedCol) {
        highestModifiedCol = rowLastCellNum;
      }
    }
    autosizeRows();
  }

  public SpreadsheetCell getCell(final int row, final int col) {
    final String address = getCellAddress(row, col);
    SpreadsheetCell cell = cells.get(address);
    if (cell == null) {
      cell = new SpreadsheetCell(this, getPoiCell(row, col));
      cells.put(address, cell);
    }
    return cell;
  }

  public SpreadsheetCell getCell(final String cellAddress) {
    final CellReference cellReference = new CellReference(cellAddress);
    return getCell(cellReference.getRow(), cellReference.getCol());
  }

  /**
   * In (1/256th of a character width)
   */
  public int getColWidth(final int col) {
    return sheet.getColumnWidth(col);
  }

  public Cell getPoiCell(final int rowNum, final int col) {
    final Row row = getPoiRow(rowNum);
    Cell cell = row.getCell(col);
    if (cell == null) {
      cell = row.createCell(col);
    }
    return cell;
  }

  @SuppressWarnings("UnusedDeclaration")
  public Sheet getPoiSheet() {
    return sheet;
  }

  /**
   * In twips (1/20th of a point)
   */
  public int getRowHeight(final int row) {
    return sheet.getRow(row).getHeight();
  }

  public void mergeCells(final int firstRow, final int lastRow, final int firstCol, final int lastCol, final Object content,
      final SpreadsheetCellStyle style) {
    setValue(firstRow, firstCol, content);
    for (int col = firstCol; col <= lastCol; col++) {
      for (int row = firstRow; row <= lastRow; row++) {
        setStyle(row, col, style);
      }
    }
    sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
  }

  public void mergeCells(final String firstCell, final String lastCell, final Object content, final SpreadsheetCellStyle style) {
    final CellReference firstReference = new CellReference(firstCell);
    final CellReference lastReference = new CellReference(lastCell);
    mergeCells(firstReference.getRow(), lastReference.getRow(), firstReference.getCol(), lastReference.getCol(), content, style);
  }

  /**
   * Returns the index of the next col after the last one written.
   */
  public int printAcross(final int row, final int col, final SpreadsheetCellStyle style, final Object... values) {
    for (int i = 0; i < values.length; i++) {
      setValue(row, col + i, values[i], style);
    }
    return col + values.length;
  }

  public void printAcross(final String cellAddress, final SpreadsheetCellStyle style, final String... values) {
    final CellReference cellReference = new CellReference(cellAddress);
    printAcross(cellReference.getRow(), cellReference.getCol(), style, values);
  }

  /**
   * Returns the index of the next row after the last one written
   */
  public int printDown(final int row, final int col, final SpreadsheetCellStyle style, final Object... values) {
    for (int i = 0; i < values.length; i++) {
      setValue(row + i, col, values[i], style);
    }
    return row + values.length;
  }

  public void printDown(final String cellAddress, final SpreadsheetCellStyle style, final String... values) {
    final CellReference cellReference = new CellReference(cellAddress);
    printDown(cellReference.getRow(), cellReference.getCol(), style, values);
  }

  public CellStyle registerStyle(final SpreadsheetCellStyle style) {
    return workbook.registerStyle(style);
  }

  public void setBottomBorder(final int row, final int firstCol, final int lastCol, final short border) {
    for (int col = firstCol; col <= lastCol; col++) {
      getCell(row, col).applyStyle(DEFAULT.withBottomBorder(border));
    }
  }

  /**
   * In (1/256th of a character width)
   */
  public void setColWidth(final int col, final int twips) {
    sheet.setColumnWidth(col, twips);
  }

  public void setLeftBorder(final int firstRow, final int lastRow, final int col, final short border) {
    for (int row = firstRow; row <= lastRow; row++) {
      getCell(row, col).applyStyle(DEFAULT.withLeftBorder(border));
    }
  }

  public void setRightBorder(final int firstRow, final int lastRow, final int col, final short border) {
    for (int row = firstRow; row <= lastRow; row++) {
      getCell(row, col).applyStyle(DEFAULT.withRightBorder(border));
    }
  }

  /**
   * In twips (1/20th of a point)
   */
  public void setRowHeight(final int row, final int twips) {
    sheet.getRow(row).setHeight((short) twips);
  }

  public void setStyle(final int firstRow, final int lastRow, final int firstCol, final int lastCol, final SpreadsheetCellStyle style) {
    for (int row = firstRow; row <= lastRow; row++) {
      for (int col = firstCol; col <= lastCol; col++) {
        getCell(row, col).setStyle(style);
      }
    }
  }

  public void setStyle(final int row, final int col, final SpreadsheetCellStyle style) {
    getCell(row, col).setStyle(style);
  }

  public void setStyle(final String cellAddress, final SpreadsheetCellStyle style) {
    final CellReference cellReference = new CellReference(cellAddress);
    setStyle(cellReference.getRow(), cellReference.getCol(), style);
  }

  public void setStyle(final String firstCell, final String lastCell, final SpreadsheetCellStyle style) {
    final CellReference firstReference = new CellReference(firstCell);
    final CellReference lastReference = new CellReference(lastCell);
    setStyle(firstReference.getRow(), lastReference.getRow(), firstReference.getCol(), lastReference.getCol(), style);
  }

  public void setSurroundBorder(final int firstRow, final int lastRow, final int firstCol, final int lastCol, final short border) {
    setTopBorder(firstRow, firstCol, lastCol, border);
    setBottomBorder(lastRow, firstCol, lastCol, border);
    setLeftBorder(firstRow, lastRow, firstCol, border);
    setRightBorder(firstRow, lastRow, lastCol, border);
  }

  public void setSurroundBorder(final String firstCell, final String lastCell, final short border) {
    final CellReference firstReference = new CellReference(firstCell);
    final CellReference lastReference = new CellReference(lastCell);
    setSurroundBorder(firstReference.getRow(), lastReference.getRow(), firstReference.getCol(), lastReference.getCol(), border);
  }

  public void setTopBorder(final int row, final int firstCol, final int lastCol, final short border) {
    for (int col = firstCol; col <= lastCol; col++) {
      getCell(row, col).applyStyle(DEFAULT.withTopBorder(border));
    }
  }

  public void setValue(final int row, final int col, final Object content) {
    setValue(row, col, content, null);
  }

  public void setValue(final int row, final int col, final Object content, final SpreadsheetCellStyle style) {
    final SpreadsheetCell cell = getCell(row, col);
    cell.setValue(content);
    if (style != null) {
      cell.setStyle(style);
    }
    recordCellModified(row, col);
  }

  public void setValue(final String cellAddress, final Object content) {
    setValue(cellAddress, content, null);
  }

  public void setValue(final String cellAddress, final Object content, final SpreadsheetCellStyle style) {
    final CellReference cellReference = new CellReference(cellAddress);
    setValue(cellReference.getRow(), cellReference.getCol(), content, style);
  }

  private Row getPoiRow(final int rowNum) {
    Row row = sheet.getRow(rowNum);
    if (row == null) {
      row = sheet.createRow(rowNum);
    }
    return row;
  }

  private void recordCellModified(final int row, final int col) {
    if (col > highestModifiedCol) {
      highestModifiedCol = col;
    }
    if (row > highestModifiedRow) {
      highestModifiedRow = row;
    }
  }

}
