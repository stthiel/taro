package taro.spreadsheet.model;

import static com.google.common.collect.Maps.newHashMap;
import static com.google.common.primitives.Shorts.checkedCast;

import java.io.IOException;
import java.io.OutputStream;
import java.util.Collections;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SpreadsheetWorkbook {

  private final Workbook workbook;
  private final Map<SpreadsheetFont, Font> fontMap = newHashMap();
  private final Map<SpreadsheetCellStyle, CellStyle> styleMap = newHashMap();

  public SpreadsheetWorkbook() {
    this(new XSSFWorkbook());
  }

  public SpreadsheetWorkbook(final Workbook workbook) {
    this.workbook = workbook;
  }

  public SpreadsheetTab createTab(final String title) {
    return new SpreadsheetTab(this, title);
  }

  public Map<SpreadsheetCellStyle, CellStyle> getCellStyles() {
    return Collections.unmodifiableMap(styleMap);
  }

  public Map<SpreadsheetFont, Font> getFonts() {
    return Collections.unmodifiableMap(fontMap);
  }

  public Workbook getPoiWorkbook() {
    return workbook;
  }

  public SpreadsheetTab getTab(final int index) {
    return new SpreadsheetTab(this, workbook.getSheetAt(index));
  }

  public SpreadsheetTab getTab(final String title) {
    return new SpreadsheetTab(this, workbook.getSheet(title));
  }

  public CellStyle registerStyle(final SpreadsheetCellStyle style) {
    CellStyle cellStyle = styleMap.get(style);
    if (cellStyle == null) {
      cellStyle = createNewStyle(style);
      styleMap.put(style, cellStyle);
    }
    return cellStyle;
  }

  public void write(final OutputStream out) throws IOException {
    workbook.write(out);
  }

  private Font createNewFont(final SpreadsheetFont font) {
    final Font poiFont = workbook.createFont();
    if (font.getBold() != null)
      // poiFont.setBold(font.getBold());
      if (font.getFontName() != null)
        poiFont.setFontName(font.getFontName());
    if (font.getFontOffset() != null)
      poiFont.setTypeOffset(checkedCast(font.getFontOffset()));
    if (font.getItalic() != null)
      poiFont.setItalic(font.getItalic());
    if (font.getUnderline() != null)
      poiFont.setUnderline(font.getUnderline() ? Font.U_SINGLE : Font.U_NONE);
    if (font.getDoubleUnderline() != null)
      poiFont.setUnderline(font.getDoubleUnderline() ? Font.U_DOUBLE : Font.U_NONE);
    if (font.getStrikeout() != null)
      poiFont.setStrikeout(font.getStrikeout());
    if (font.getFontSizeInPoints() != null)
      poiFont.setFontHeightInPoints(checkedCast(font.getFontSizeInPoints()));
    return poiFont;
  }

  private CellStyle createNewStyle(final SpreadsheetCellStyle style) {
    final CellStyle cellStyle = workbook.createCellStyle();
    if (style.getAlign() != null)
      cellStyle.setAlignment(style.getAlign());
    if (style.getVerticalAlign() != null)
      cellStyle.setVerticalAlignment(style.getVerticalAlign());
    if (style.getTopBorder() != null)
      cellStyle.setBorderTop(style.getTopBorder());
    if (style.getLeftBorder() != null)
      cellStyle.setBorderLeft(style.getLeftBorder());
    if (style.getBottomBorder() != null)
      cellStyle.setBorderBottom(style.getBottomBorder());
    if (style.getRightBorder() != null)
      cellStyle.setBorderRight(style.getRightBorder());
    if (style.getLocked() != null)
      cellStyle.setLocked(style.getLocked());
    if (style.isHidden() != null)
      cellStyle.setHidden(style.isHidden());
    if (style.getWrapText() != null)
      cellStyle.setWrapText(style.getWrapText());
    if (style.getIndention() != null)
      cellStyle.setIndention(checkedCast(style.getIndention()));
    if (style.getRotation() != null)
      cellStyle.setRotation(checkedCast(style.getRotation()));
    // if (style.getTopBorderColor() != null)
    // cellStyle.setTopBorderColor(new XSSFColor(style.getTopBorderColor()));
    // if (style.getLeftBorderColor() != null)
    // cellStyle.setLeftBorderColor(new XSSFColor(style.getLeftBorderColor()));
    // if (style.getBottomBorderColor() != null)
    // cellStyle.setBottomBorderColor(new XSSFColor(style.getBottomBorderColor()));
    // if (style.getRightBorderColor() != null)
    // cellStyle.setRightBorderColor(new XSSFColor(style.getRightBorderColor()));

    if (style.getFont() != null) {
      cellStyle.setFont(registerFont(style.getFont()));
    }

    if (style.getBackgroundColor() != null) {
      cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
      // cellStyle.setFillForegroundColor(new XSSFColor(style.getBackgroundColor()));
    }

    if (style.getDataFormatString() != null) {
      cellStyle.setDataFormat(workbook.createDataFormat().getFormat(style.getDataFormatString()));
    }

    return cellStyle;
  }

  private Font registerFont(final SpreadsheetFont font) {
    Font poiFont = fontMap.get(font);
    if (poiFont == null) {
      poiFont = createNewFont(font);
      fontMap.put(font, poiFont);
    }
    return poiFont;
  }

}
