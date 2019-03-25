package app;

import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import cronapi.CronapiMetaData;
import cronapi.ParamMetaData;
import cronapi.Var;
import cronapi.CronapiMetaData.ObjectType;

/**
 * Utilitário para manipulação de arquivos excel ...
 * 
 * @author Ricardo Caldas
 * @version 1.0
 * @since 2019-01-31
 *
 */

@CronapiMetaData(categoryName = "ExcelUtils")
public class ManipularExcel {

	@CronapiMetaData(type = "function", name = "{{createSpreadSheet}}", description = "{{createSpreadSheetDescription}}", returnType = ObjectType.OBJECT)
	public static Var createSpreadSheet() throws Exception {
		XSSFWorkbook workbook = new XSSFWorkbook();
		return Var.valueOf(workbook);
	}

	@CronapiMetaData(type = "function", name = "{{createSheet}}", description = "{{createSheetDescription}}", returnType = ObjectType.OBJECT)
	public static Var createSheet(
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramSpreadSheet}}") Var workbook,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramSheetName}}") Var name)
			throws Exception {
		Sheet sheet = workbook.getTypedObject(XSSFWorkbook.class)
				.createSheet(name.toString());
		return Var.valueOf(sheet);
	}

	@CronapiMetaData(type = "function", name = "{{createLine}}", description = "{{createLineDescription}}", returnType = ObjectType.UNKNOWN)
	public static Var createLine(
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramRowNum}}") Var rowNumber,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramSheet}}") Var sheet)
			throws Exception {
		Row row = sheet.getTypedObject(XSSFSheet.class).createRow(rowNumber.getObjectAsInt());
		return Var.valueOf(row);
	}
	
	@CronapiMetaData(type = "function", name = "{{getCellValue}}", description = "{{getCellValueDescription}}", returnType = ObjectType.OBJECT)
	public static Var getCellValue(
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramRowNum}}") Var rowNumber,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramColNum}}") Var columnNumber,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramSheet}}") Var sheet)
			throws Exception {

		XSSFRow row = sheet.getTypedObject(XSSFSheet.class).getRow(rowNumber.getObjectAsInt());
		
		XSSFCell cell = row.getCell(columnNumber.getObjectAsInt());
		Var valor;
		
		switch (cell.getCellType()) {
		case XSSFCell.CELL_TYPE_BLANK:
			valor = Var.valueOf(cell.getRawValue());
			break;
		case XSSFCell.CELL_TYPE_BOOLEAN:
			valor = Var.valueOf(cell.getBooleanCellValue());
			break;
		case XSSFCell.CELL_TYPE_ERROR:
			valor = Var.valueOf(cell.getErrorCellString());
			break;
		case XSSFCell.CELL_TYPE_FORMULA:
			valor = Var.valueOf(cell.getCellFormula());
			break;
		case XSSFCell.CELL_TYPE_NUMERIC:
			valor = Var.valueOf(cell.getNumericCellValue());
			break;
		case XSSFCell.CELL_TYPE_STRING:
			valor = Var.valueOf(cell.getStringCellValue());
			break;
		default:
		  valor = Var.valueOf(cell.getStringCellValue());
		}
		
  return valor;
	}

	@CronapiMetaData(type = "function", name = "{{insertCellValue}}", description = "{{insertCellValueDescription}}", returnType = ObjectType.VOID)
	public static void insertCellValue(
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramRowNum}}") Var rowNumber,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramColNum}}") Var columnNumber,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramSheet}}") Var sheet,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramColumnValue}}") Var value)
			throws Exception {

		Row row = sheet.getTypedObject(XSSFSheet.class).getRow(rowNumber.getObjectAsInt());
		if (row == null) {
			row = sheet.getTypedObject(XSSFSheet.class).createRow(rowNumber.getObjectAsInt());
		}

		Cell cell = row.getCell(columnNumber.getObjectAsInt());
		if (cell == null) {
			cell = row.createCell(columnNumber.getObjectAsInt());
		}
		
     switch (value.getType()) {
      case STRING:
        cell.setCellValue(value.getObjectAsString());
        break;
      case INT:
          cell.setCellValue(value.getObjectAsInt());
        break;
      case DOUBLE:
          cell.setCellValue(value.getObjectAsDouble());
        break;
      case BOOLEAN:
          cell.setCellValue(value.getObjectAsBoolean());
        break;
      case DATETIME:
          cell.setCellValue(value.getObjectAsDateTime());
        break;
      default:
        cell.setCellValue(value.getObjectAsString());
    }
	}

	@CronapiMetaData(type = "function", name = "{{setCellType}}", description = "{{setCellTypeDescription}}", returnType = ObjectType.VOID)
	public static void setCellType(
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramRowNum}}") Var rowNumber,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramColNum}}") Var columnNumber,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramSheet}}") Var sheet,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramCellType}}", blockType = "util_dropdown", keys = {
					"{{optionBlank}}", "{{optionBoolean}}", "{{optionError}}", "{{optionFormula}}", "{{optionNumeric}}", "{{optionText}}" }) Var cellType)
			throws Exception {

		Row row = sheet.getTypedObject(XSSFSheet.class).getRow(rowNumber.getObjectAsInt());
		if (row == null) {
			row = sheet.getTypedObject(XSSFSheet.class).createRow(rowNumber.getObjectAsInt());
		}

		Cell cell = row.getCell(columnNumber.getObjectAsInt());
		if (cell == null) {
			cell = row.createCell(columnNumber.getObjectAsInt());
		}

		switch (cellType.getObjectAsString()) {
		case "{{optionBlank}}":
			cell.setCellType(XSSFCell.CELL_TYPE_BLANK);
			break;
		case "{{optionBoolean}}":
			cell.setCellType(XSSFCell.CELL_TYPE_BOOLEAN);
			break;
		case "{{optionError}}":
			cell.setCellType(XSSFCell.CELL_TYPE_ERROR);
			break;
		case "{{optionFormula}}":
			cell.setCellType(XSSFCell.CELL_TYPE_FORMULA);
			break;
		case "{{optionNumeric}}":
			cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
			break;
		case "{{optionText}}":
			cell.setCellType(XSSFCell.CELL_TYPE_STRING);
			break;
		}
	}

	@CronapiMetaData(type = "function", name = "{{alignCellTextHorizontal}}", description = "{{alignCellTextHorizontalDescription}}", returnType = ObjectType.VOID)
	public static void alignCellTextHorizontal(
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramRowNum}}") Var rowNumber,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramColNum}}") Var columnNumber,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramSheet}}") Var sheet,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramHorizontalCellAlignment}}", blockType = "util_dropdown", keys = {
					"{{optionCenter}}", "{{optionLeft}}", "{{optionRight}}" }) Var cellType)
			throws Exception {

		Row row = sheet.getTypedObject(XSSFSheet.class).getRow(rowNumber.getObjectAsInt());
		if (row == null) {
			row = sheet.getTypedObject(XSSFSheet.class).createRow(rowNumber.getObjectAsInt());
		}

		Cell cell = row.getCell(columnNumber.getObjectAsInt());
		if (cell == null) {
			cell = row.createCell(columnNumber.getObjectAsInt());
		}

		CellStyle style = sheet.getTypedObject(XSSFSheet.class).getWorkbook().createCellStyle();

		switch (cellType.getObjectAsString()) {
		case "{{optionCenter}}":
			style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			break;
		case "{{optionLeft}}":
			style.setAlignment(XSSFCellStyle.ALIGN_LEFT);
			break;
		case "{{optionRight}}":
			style.setAlignment(XSSFCellStyle.ALIGN_RIGHT);
			break;
		}
		cell.setCellStyle(style);
	}

	@CronapiMetaData(type = "function", name = "{{alignCellTextVertical}}", description = "{{alignCellTextVerticalDescription}}", returnType = ObjectType.VOID)
	public static void alignCellTextVertical(
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramRowNum}}") Var rowNumber,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramColNum}}") Var columnNumber,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramSheet}}") Var sheet,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramVerticalCellAligment}}", blockType = "util_dropdown", keys = {
					"{{optionBottom}}", "{{optionCenter}}", "{{optionJustify}}", "{{optionTop}}" }) Var cellType)
			throws Exception {

		Row row = sheet.getTypedObject(XSSFSheet.class).getRow(rowNumber.getObjectAsInt());
		if (row == null) {
			row = sheet.getTypedObject(XSSFSheet.class).createRow(rowNumber.getObjectAsInt());
		}

		Cell cell = row.getCell(columnNumber.getObjectAsInt());
		if (cell == null) {
			cell = row.createCell(columnNumber.getObjectAsInt());
		}

		CellStyle style = sheet.getTypedObject(XSSFSheet.class).getWorkbook().createCellStyle();

		switch (cellType.getObjectAsString()) {
		case "{{optionBottom}}":
			style.setVerticalAlignment(XSSFCellStyle.VERTICAL_BOTTOM);
			break;
		case "{{optionCenter}}":
			style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			break;
		case "{{optionJustify}}":
			style.setVerticalAlignment(XSSFCellStyle.VERTICAL_JUSTIFY);
			break;
		case "{{optionTop}}":
			style.setVerticalAlignment(XSSFCellStyle.VERTICAL_TOP);
			break;
		}
		cell.setCellStyle(style);
	}

	@CronapiMetaData(type = "function", name = "{{FontProperties}}", description = "{{FontPropertiesDescription}}", returnType = ObjectType.VOID)
	public static void setFontProperties(
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramRowNum}}") Var rowNumber,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramColNum}}") Var columnNumber,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramSheet}}") Var sheet,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramFontType}}", blockType = "util_dropdown", keys = {
					"{{optionItalic}}", "{{optionBold}}", "{{optionUnderline}}", "{{optionStrikeout}}" }) Var fontType,
			@ParamMetaData(type = ObjectType.STRING, description = "{{fontColor}}", blockType = "util_dropdown", keys = {
					"{{colorYellow}}", "{{colorWhite}}", "{{colorRed}}", "{{colorPink}}", "{{colorGreen}}", "{{colorOrange}}", "{{colorMagenta}}", "{{colorLightGrey}}", "{{colorGreen}}",
					"{{colorGrey}}", "{{colorDarkGrey}}", "{{colorCyan}}", "{{colorBlue}}", "{{colorBlack}}" }) Var fontColor)
			throws Exception {

		Row row = sheet.getTypedObject(XSSFSheet.class).getRow(rowNumber.getObjectAsInt());
		if (row == null) {
			row = sheet.getTypedObject(XSSFSheet.class).createRow(rowNumber.getObjectAsInt());
		}

		Cell cell = row.getCell(columnNumber.getObjectAsInt());
		if (cell == null) {
			cell = row.createCell(columnNumber.getObjectAsInt());
		}
		XSSFWorkbook workbook = sheet.getTypedObject(XSSFSheet.class).getWorkbook();
		CellStyle style = workbook.createCellStyle();
		XSSFFont font = workbook.createFont();
		XSSFColor color;

		switch (fontType.getObjectAsString()) {
		case "{{optionItalic}}":
			font.setItalic(true);
			break;
		case "{{optionBold}}":
			font.setBold(true);
			break;
		case "{{optionUnderscore}}":
			font.setUnderline(XSSFFont.U_SINGLE);
			break;
		case "{{optionStrikeout}}":
			font.setStrikeout(true);
			break;
		}

		switch (fontColor.getObjectAsString()) {
		case "{{colorYellow}}":
			color = new XSSFColor(Color.YELLOW);
			font.setColor(color);
			break;
		case "{{colorWhite}}":
			color = new XSSFColor(Color.WHITE);
			font.setColor(color);
			break;
		case "{{colorRed}}":
			color = new XSSFColor(Color.RED);
			font.setColor(color);
			break;
		case "{{colorPink}}":
			color = new XSSFColor(Color.PINK);
			font.setColor(color);
			break;
		case "{{colorOrange}}":
			color = new XSSFColor(Color.ORANGE);
			font.setColor(color);
			break;
		case "{{colorMagenta}}":
			color = new XSSFColor(Color.MAGENTA);
			font.setColor(color);
			break;
		case "{{colorLightGrey}}":
			color = new XSSFColor(Color.LIGHT_GRAY);
			font.setColor(color);
			break;
		case "{{colorGreen}}":
			color = new XSSFColor(Color.GREEN);
			font.setColor(color);
			break;
		case "{{colorGrey}}":
			color = new XSSFColor(Color.GRAY);
			font.setColor(color);
			break;
		case "{{colorDarkGrey}}":
			color = new XSSFColor(Color.DARK_GRAY);
			font.setColor(color);
			break;
		case "{{colorCyan}}":
			color = new XSSFColor(Color.CYAN);
			font.setColor(color);
			break;
		case "{{colorBlue}}":
			color = new XSSFColor(Color.BLUE);
			font.setColor(color);
			break;
		case "{{colorBlack}}":
			color = new XSSFColor(Color.BLACK);
			font.setColor(color);
			break;
		}
		style.setFont(font);
		cell.setCellStyle(style);
	}

	@CronapiMetaData(type = "function", name = "{{setCellBorder}}", description = "{{setCellBorderDescription}}", returnType = ObjectType.VOID)
	public static void setBorder(
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramRowNum}}") Var rowNumber,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramColNum}}") Var columnNumber,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramSheet}}") Var sheet,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramBorderType}}", blockType = "util_dropdown", keys = {
					"{{optionDotted}}", "{{optionDashed}}", "{{optionThick}}", "{{optionThin}}" }) Var borderType)
			throws Exception {

		Row row = sheet.getTypedObject(XSSFSheet.class).getRow(rowNumber.getObjectAsInt());
		if (row == null) {
			row = sheet.getTypedObject(XSSFSheet.class).createRow(rowNumber.getObjectAsInt());
		}

		Cell cell = row.getCell(columnNumber.getObjectAsInt());
		if (cell == null) {
			cell = row.createCell(columnNumber.getObjectAsInt());
		}

		CellStyle style = sheet.getTypedObject(XSSFSheet.class).getWorkbook().createCellStyle();

		switch (borderType.getObjectAsString()) {
		case "{{optionDotted}}":
			style.setBorderRight(XSSFCellStyle.BORDER_DOTTED);
			style.setBorderLeft(XSSFCellStyle.BORDER_DOTTED);
			style.setBorderTop(XSSFCellStyle.BORDER_DOTTED);
			style.setBorderBottom(XSSFCellStyle.BORDER_DOTTED);
			break;
		case "{{optionDashed}}":
			style.setBorderRight(XSSFCellStyle.BORDER_DASHED);
			style.setBorderLeft(XSSFCellStyle.BORDER_DASHED);
			style.setBorderTop(XSSFCellStyle.BORDER_DASHED);
			style.setBorderBottom(XSSFCellStyle.BORDER_DASHED);
			break;
		case "{{optionThick}}":
			style.setBorderRight(XSSFCellStyle.BORDER_THICK);
			style.setBorderLeft(XSSFCellStyle.BORDER_THICK);
			style.setBorderTop(XSSFCellStyle.BORDER_THICK);
			style.setBorderBottom(XSSFCellStyle.BORDER_THICK);
			break;
		case "{{optionThin}}":
			style.setBorderRight(XSSFCellStyle.BORDER_THIN);
			style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			style.setBorderTop(XSSFCellStyle.BORDER_THIN);
			style.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			break;
		}

		cell.setCellStyle(style);
	}

	@CronapiMetaData(type = "function", name = "{{setBorderColor}}", description = "{{setBorderColorDescription}}", returnType = ObjectType.VOID)
	public static void setBorderColor(
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramRowNum}}") Var rowNumber,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramColNum}}") Var columnNumber,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramSheet}}") Var sheet,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramColor}}", blockType = "util_dropdown", keys = {
					"{{colorYellow}}", "{{colorWhite}}", "{{colorRed}}", "{{colorPink}}", "{{colorGreen}}", "{{colorOrange}}", "{{colorMagenta}}", "{{colorLightGrey}}", "{{colorGreen}}",
					"{{colorGrey}}", "{{colorDarkGrey}}", "{{colorCyan}}", "{{colorBlue}}", "{{colorBlack}}" }) Var cellType)
			throws Exception {

		Row row = sheet.getTypedObject(XSSFSheet.class).getRow(rowNumber.getObjectAsInt());
		if (row == null) {
			row = sheet.getTypedObject(XSSFSheet.class).createRow(rowNumber.getObjectAsInt());
		}

		Cell cell = row.getCell(columnNumber.getObjectAsInt());
		if (cell == null) {
			cell = row.createCell(columnNumber.getObjectAsInt());
		}

		XSSFCellStyle style = sheet.getTypedObject(XSSFSheet.class).getWorkbook().createCellStyle();
		XSSFColor color;

		switch (cellType.getObjectAsString()) {
		case "{{colorYellow}}":
			color = new XSSFColor(Color.YELLOW);
			style.setTopBorderColor(color);
			style.setRightBorderColor(color);
			style.setLeftBorderColor(color);
			style.setBottomBorderColor(color);
			break;
		case "{{colorWhite}}":
			color = new XSSFColor(Color.WHITE);
			style.setTopBorderColor(color);
			style.setRightBorderColor(color);
			style.setLeftBorderColor(color);
			style.setBottomBorderColor(color);
			break;
		case "{{colorRed}}":
			color = new XSSFColor(Color.RED);
			style.setTopBorderColor(color);
			style.setRightBorderColor(color);
			style.setLeftBorderColor(color);
			style.setBottomBorderColor(color);
			break;
		case "{{colorPink}}":
			color = new XSSFColor(Color.PINK);
			style.setTopBorderColor(color);
			style.setRightBorderColor(color);
			style.setLeftBorderColor(color);
			style.setBottomBorderColor(color);
			break;
		case "{{colorOrange}}":
			color = new XSSFColor(Color.ORANGE);
			style.setTopBorderColor(color);
			style.setRightBorderColor(color);
			style.setLeftBorderColor(color);
			style.setBottomBorderColor(color);
			break;
		case "{{colorMagenta}}":
			color = new XSSFColor(Color.MAGENTA);
			style.setTopBorderColor(color);
			style.setRightBorderColor(color);
			style.setLeftBorderColor(color);
			style.setBottomBorderColor(color);
			break;
		case "{{colorLightGrey}}":
			color = new XSSFColor(Color.LIGHT_GRAY);
			style.setTopBorderColor(color);
			style.setRightBorderColor(color);
			style.setLeftBorderColor(color);
			style.setBottomBorderColor(color);
			break;
		case "{{colorGreen}}":
			color = new XSSFColor(Color.GREEN);
			style.setTopBorderColor(color);
			style.setRightBorderColor(color);
			style.setLeftBorderColor(color);
			style.setBottomBorderColor(color);
			break;
		case "{{colorGrey}}":
			color = new XSSFColor(Color.GRAY);
			style.setTopBorderColor(color);
			style.setRightBorderColor(color);
			style.setLeftBorderColor(color);
			style.setBottomBorderColor(color);
			break;
		case "{{colorDarkGrey}}":
			color = new XSSFColor(Color.DARK_GRAY);
			style.setTopBorderColor(color);
			style.setRightBorderColor(color);
			style.setLeftBorderColor(color);
			style.setBottomBorderColor(color);
			break;
		case "{{colorCyan}}":
			color = new XSSFColor(Color.CYAN);
			style.setTopBorderColor(color);
			style.setRightBorderColor(color);
			style.setLeftBorderColor(color);
			style.setBottomBorderColor(color);
			break;
		case "{{colorBlue}}":
			color = new XSSFColor(Color.BLUE);
			style.setTopBorderColor(color);
			style.setRightBorderColor(color);
			style.setLeftBorderColor(color);
			style.setBottomBorderColor(color);
			break;
		case "{{colorBlack}}":
			color = new XSSFColor(Color.BLACK);
			style.setTopBorderColor(color);
			style.setRightBorderColor(color);
			style.setLeftBorderColor(color);
			style.setBottomBorderColor(color);
			break;
		}
		cell.setCellStyle(style);
	}

	@CronapiMetaData(type = "function", name = "{{setBackgroundColor}}", description = "{{setBackgroundColorDescription}}", returnType = ObjectType.VOID)
	public static void setBackgroundColor(
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramRowNum}}") Var rowNumber,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramColNum}}") Var columnNumber,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramSheet}}") Var sheet,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramCellColor}}", blockType = "util_dropdown", keys = {
					"{{colorYellow}}", "{{colorWhite}}", "{{colorRed}}", "{{colorPink}}", "{{colorGreen}}", "{{colorOrange}}", "{{colorMagenta}}", "{{colorLightGrey}}", "{{colorGreen}}",
					"{{colorGrey}}", "{{colorDarkGrey}}", "{{colorCyan}}", "{{colorBlue}}", "{{colorBlack}}" }) Var cellType)
			throws Exception {

		Row row = sheet.getTypedObject(XSSFSheet.class).getRow(rowNumber.getObjectAsInt());
		if (row == null) {
			row = sheet.getTypedObject(XSSFSheet.class).createRow(rowNumber.getObjectAsInt());
		}

		Cell cell = row.getCell(columnNumber.getObjectAsInt());
		if (cell == null) {
			cell = row.createCell(columnNumber.getObjectAsInt());
		}

		XSSFCellStyle style = sheet.getTypedObject(XSSFSheet.class).getWorkbook().createCellStyle();
		XSSFColor color;

		switch (cellType.getObjectAsString()) {
		case "{{colorYellow}}":
			color = new XSSFColor(Color.YELLOW);
			style.setFillForegroundColor(color);
			break;
		case "{{colorWhite}}":
			color = new XSSFColor(Color.WHITE);
			style.setFillForegroundColor(color);
			break;
		case "{{colorRed}}":
			color = new XSSFColor(Color.RED);
			style.setFillForegroundColor(color);
			break;
		case "{{colorPink}}":
			color = new XSSFColor(Color.PINK);
			style.setFillForegroundColor(color);
			break;
		case "{{colorOrange}}":
			color = new XSSFColor(Color.ORANGE);
			style.setFillForegroundColor(color);
			break;
		case "{{colorMagenta}}":
			color = new XSSFColor(Color.MAGENTA);
			style.setFillForegroundColor(color);
			break;
		case "{{colorLightGrey}}":
			color = new XSSFColor(Color.LIGHT_GRAY);
			style.setFillForegroundColor(color);
			break;
		case "{{colorGreen}}":
			color = new XSSFColor(Color.GREEN);
			style.setFillForegroundColor(color);
			break;
		case "{{colorGrey}}":
			color = new XSSFColor(Color.GRAY);
			style.setFillForegroundColor(color);
			break;
		case "{{colorDarkGrey}}":
			color = new XSSFColor(Color.DARK_GRAY);
			style.setFillForegroundColor(color);
			break;
		case "{{colorCyan}}":
			color = new XSSFColor(Color.CYAN);
			style.setFillForegroundColor(color);
			break;
		case "{{colorBlue}}":
			color = new XSSFColor(Color.BLUE);
			style.setFillForegroundColor(color);
			break;
		case "{{colorBlack}}":
			color = new XSSFColor(Color.BLACK);
			style.setFillForegroundColor(color);
			break;
		}
		cell.setCellStyle(style);
	}

	@CronapiMetaData(type = "function", name = "{{createLineWithValue}}", description = "{{createLineWithValueDescription}}", returnType = ObjectType.UNKNOWN)
	public static void createLineWithValueSet(
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramSheet}}") Var sheet,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramRowNum}}") Var rowNumber,
			@ParamMetaData(type = ObjectType.OBJECT, description = "{{paramValues}}") Var name)
			throws Exception {
		Row row = sheet.getTypedObject(XSSFSheet.class).createRow(rowNumber.getObjectAsInt());
		for (int i = 0; i < name.getObjectAsList().size(); i++) {
			Cell cell = row.createCell(i);
			switch (name.getObjectAsList().get(i).getType()) {
      case STRING:
        cell.setCellValue(name.getObjectAsList().get(i).getObjectAsString());
        break;
      case INT:
          cell.setCellValue(name.getObjectAsList().get(i).getObjectAsInt());
        break;
      case DOUBLE:
          cell.setCellValue(name.getObjectAsList().get(i).getObjectAsDouble());
        break;
      case BOOLEAN:
          cell.setCellValue(name.getObjectAsList().get(i).getObjectAsBoolean());
        break;
      case DATETIME:
          cell.setCellValue(name.getObjectAsList().get(i).getObjectAsDateTime());
        break;
      default:
        cell.setCellValue(name.getObjectAsList().get(i).getObjectAsString());
    }
		}
	}

	@CronapiMetaData(type = "function", name = "{{createHyperlink}}", description = "{{createHyperlinkDescription}}", returnType = ObjectType.VOID)
	public static void createHyperlink(
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramSheet}}") Var sheet,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramRowNum}}") Var rowNumber,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramColNum}}") Var columnNumber,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramHyperlinkAddress}}") Var address,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramHyperlinkType}}", blockType = "util_dropdown", keys = {
					"{{optionDocument}}", "{{optionEmail}}", "{{optionFile}}", "{{optionWebUrl}}" }) Var hyperlinkType)
			throws Exception {

		XSSFWorkbook workbook = sheet.getTypedObject(XSSFSheet.class).getWorkbook();

		XSSFCellStyle hlinkstyle = workbook.createCellStyle();
		XSSFFont hlinkfont = workbook.createFont();
		hlinkfont.setUnderline(XSSFFont.U_SINGLE);
		hlinkfont.setColor(new XSSFColor(Color.BLUE));
		hlinkstyle.setFont(hlinkfont);

		Row row = sheet.getTypedObject(XSSFSheet.class).getRow(rowNumber.getObjectAsInt());
		if (row == null)
			row = sheet.getTypedObject(XSSFSheet.class).createRow(rowNumber.getObjectAsInt());

		Cell cell = row.getCell(columnNumber.getObjectAsInt());
		if (cell == null)
			cell = row.createCell(columnNumber.getObjectAsInt());

		cell.setCellValue(address.getObjectAsString());

		XSSFCreationHelper helper = workbook.getCreationHelper();
		XSSFHyperlink hyperlink;

		switch (hyperlinkType.getObjectAsString()) {
		case "{{optionDocument}}":
			hyperlink = (XSSFHyperlink) helper.createHyperlink(Hyperlink.LINK_DOCUMENT);
			cell.setHyperlink(hyperlink);
			break;
		case "{{optionEmail}}":
			hyperlink = (XSSFHyperlink) helper.createHyperlink(Hyperlink.LINK_EMAIL);
			cell.setHyperlink(hyperlink);
			break;
		case "{{optionFile}}":
			hyperlink = (XSSFHyperlink) helper.createHyperlink(Hyperlink.LINK_FILE);
			cell.setHyperlink(hyperlink);
			break;
		case "{{optionWebUrl}}":
			hyperlink = (XSSFHyperlink) helper.createHyperlink(Hyperlink.LINK_URL);
			cell.setHyperlink(hyperlink);
			break;
		}
		cell.setCellStyle(hlinkstyle);
	}

	@CronapiMetaData(type = "function", name = "{{mergeCells}}", description = "{{mergeCellsDescription}}", returnType = ObjectType.VOID)
	public static void mergeCells(
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramSheet}}") Var sheet,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramFirstRow}}") Var firstRow,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramLastRow}}") Var lastRow,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramFirstCol}}") Var firstCol,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramLastCol}}") Var lastCol)
			throws Exception {
		CellRangeAddress range = new CellRangeAddress(firstRow.getObjectAsInt(),
				lastRow.getObjectAsInt(), firstCol.getObjectAsInt(),
				lastCol.getObjectAsInt());
		sheet.getTypedObject(XSSFSheet.class).addMergedRegion(range);
	}

	@CronapiMetaData(type = "function", name = "{{autoSizeColumn}}", description = "{{autoSizeColumnDescription}}", returnType = ObjectType.VOID)
	public static void autoSize(
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramSheet}}") Var sheet,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramColNum}}") Var column)
			throws Exception {
		sheet.getTypedObject(XSSFSheet.class).autoSizeColumn(column.getObjectAsInt());
	}

	@CronapiMetaData(type = "function", name = "{{saveExcelFile}}", description = "{{saveExcelFileDescription}}", returnType = ObjectType.VOID)
	public static void saveSpreadSheet(
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramSpreadSheet}}") Var workbook,
			@ParamMetaData(type = ObjectType.STRING, description = "{{paramSavePath}}") Var path)
			throws Exception {
		File arquivo = new File(path.toString());
		FileOutputStream out = new FileOutputStream(arquivo);
		workbook.getTypedObject(XSSFWorkbook.class).write(out);
		out.close();
	}

}