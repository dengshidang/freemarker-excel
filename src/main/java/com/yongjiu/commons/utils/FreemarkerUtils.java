package com.yongjiu.commons.utils;

import com.yongjiu.dto.freemarker.input.ExcelImageInput;
import com.yongjiu.dto.freemarker.input.FreemarkerInput;
import com.yongjiu.entity.excel.Cell;
import com.yongjiu.entity.excel.Row;
import com.yongjiu.entity.excel.Table;
import com.yongjiu.entity.excel.*;
import com.yongjiu.util.ColorUtil;
import freemarker.template.*;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.*;
import org.dom4j.Document;
import org.dom4j.io.SAXReader;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.CollectionUtils;
import org.springframework.util.ObjectUtils;

import javax.imageio.ImageIO;
import javax.servlet.http.HttpServletResponse;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * author 大脑补丁
 * project freemarker-excel
 * description: freemarker工具类
 * create 2020-04-14 09:43
 */
public class FreemarkerUtils {

	private static final Logger log = LoggerFactory.getLogger(FreemarkerUtils.class);

	/**
	 * 导出Excel到指定文件中
	 *
	 * @param dataMap          数据源
	 * @param templateName     模板名称（包含文件后缀名.ftl）
	 * @param templateFilePath 模板所在路径（不能为空，当前路径传空字符：""）
	 * @param fileFullPath     文件完整路径（如：usr/local/fileName.xls）
	 * @author XuChao on 2021-04-05 11:51
	 */
	@SuppressWarnings("rawtypes")
	public static void exportToFile(Map dataMap, String templateName, String templateFilePath, String fileFullPath) {
		try {
			File file = new File(fileFullPath);
			FileUtils.forceMkdirParent(file);
			FileOutputStream outputStream = new FileOutputStream(file);
			exportToStream(dataMap, templateName, templateFilePath, outputStream);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * 导出Excel到输出流
	 *
	 * @param dataMap          数据源
	 * @param templateName     模板名称（包含文件后缀名.ftl）
	 * @param templateFilePath 模板所在路径（不能为空，当前路径传空字符：""）
	 * @param outputStream    输出流
	 * @author `XuChao` on 2021-04-05 11:52
	 */
	@SuppressWarnings("rawtypes")
	public static void exportToStream(Map dataMap, String templateName, String templateFilePath,
			FileOutputStream outputStream) {
		try {
			Template template = getTemplate(templateName, templateFilePath);
			OutputStreamWriter outputWriter = new OutputStreamWriter(outputStream, "UTF-8");
			Writer writer = new BufferedWriter(outputWriter);
			template.process(dataMap, writer);
			writer.flush();
			writer.close();
			outputStream.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * 导出到文件中（导出到硬盘，xls格式）
	 *
	 * param excelFilePath
	 * param freemarkerInput
	 * author XuChao on 2021-04-14 15:34
	 */
	public static void exportImageExcel(String excelFilePath, FreemarkerInput freemarkerInput) {
		try {
			File file = new File(excelFilePath);
			FileUtils.forceMkdirParent(file);
			FileOutputStream outputStream = new FileOutputStream(file);
			createImageExcleToStream(freemarkerInput, outputStream);
			// 删除xml缓存文件
			FileUtils.forceDelete(new File(freemarkerInput.getXmlTempFile() + freemarkerInput.getFileName() + ".xml"));
			log.info("导出成功,导出到目录：" + file.getCanonicalPath());
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	/**
	 * 导出到文件中（导出到硬盘，xlsx格式）
	 *
	 * param excelFilePath
	 * param freemarkerInput
	 * author XuChao on 2021-06-25 23:34
	 */
	public static void exportImageExcelNew(String excelFilePath, FreemarkerInput freemarkerInput) {
		try {
			File file = new File(excelFilePath);
			FileUtils.forceMkdirParent(file);
			FileOutputStream outputStream = new FileOutputStream(file);
			createExcelToStream(freemarkerInput, outputStream);
			// 删除xml缓存文件
			FileUtils.forceDelete(new File(freemarkerInput.getXmlTempFile() + freemarkerInput.getFileName() + ".xml"));
			log.info("导出成功,导出到目录：" + file.getCanonicalPath());
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	/**
	 * 导出到response输出流中(用于浏览器调用接口,支持Excel2007版，xlsx格式)
	 *
	 * param response
	 * param freemarkerInput
	 */
	public static void exportImageExcelNew(HttpServletResponse response, FreemarkerInput freemarkerInput) {
		try {
			OutputStream outputStream = response.getOutputStream();
			// 写入excel文件
			response.reset();
			response.setContentType("application/msexcel;charset=UTF-8");
			response.setHeader("Content-Disposition",
					"attachment;filename=\"" + new String((freemarkerInput.getFileName() + ".xlsx").getBytes("GBK"),
							"ISO8859-1") + "\"");
			response.setHeader("Response-Type", "Download");
			createExcelToStream(freemarkerInput, outputStream);
			// 删除xml缓存文件
			FileUtils.forceDelete(new File(freemarkerInput.getXmlTempFile() + freemarkerInput.getFileName() + ".xml"));
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * 导出到response输出流中(用于浏览器调用接口,支持Excel2003版，xls格式)
	 *
	 * param response
	 * param freemarkerInput
	 * author XuChao on 2021-06-25 23:34
	 */
	public static void exportImageExcel(HttpServletResponse response, FreemarkerInput freemarkerInput) {
		try {
			OutputStream outputStream = response.getOutputStream();
			// 写入excel文件
			response.reset();
			response.setContentType("application/msexcel;charset=UTF-8");
			response.setHeader("Content-Disposition",
					"attachment;filename=\"" + new String((freemarkerInput.getFileName() + ".xls").getBytes("GBK"),
							"ISO8859-1") + "\"");
			response.setHeader("Response-Type", "Download");
			createImageExcleToStream(freemarkerInput, outputStream);
			// 删除xml缓存文件
			FileUtils.forceDelete(new File(freemarkerInput.getXmlTempFile() + freemarkerInput.getFileName() + ".xml"));
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// 获取项目templates文件夹下的模板
	private static Template getTemplate(String templateName, String filePath) throws IOException {
		Configuration configuration = new Configuration(Configuration.VERSION_2_3_28);
		configuration.setDefaultEncoding("UTF-8");
		configuration.setTemplateUpdateDelayMilliseconds(0);
		configuration.setEncoding(Locale.CHINA, "UTF-8");
		configuration.setTemplateExceptionHandler(TemplateExceptionHandler.RETHROW_HANDLER);
		configuration.setClassForTemplateLoading(FreemarkerUtils.class,  filePath);
		configuration.setOutputEncoding("UTF-8");
		return configuration.getTemplate(templateName, "UTF-8");
	}

	/**
	 * 导出Excel到输出流（支持Excel2003版，xls格式）
	 *
	 * @param freemarkerInput
	 * @param outputStream
	 */
	private static void createImageExcleToStream(FreemarkerInput freemarkerInput, OutputStream outputStream) {
		Writer out = null;
		try {
			// 创建xml文件
			Template template = getTemplate(freemarkerInput.getTemplateName(), freemarkerInput.getTemplateFilePath());
			File tempXMLFile = new File(freemarkerInput.getXmlTempFile() + freemarkerInput.getFileName() + ".xml");
			FileUtils.forceMkdirParent(tempXMLFile);
			out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(tempXMLFile), "UTF-8"));
			template.process(freemarkerInput.getDataMap(), out);
			if (log.isDebugEnabled()) {
				log.debug("1.完成将文本数据导入到XML文件中");
			}
			SAXReader reader = new SAXReader();
			Document document = reader.read(tempXMLFile);
			Map<String, Style> styleMap = readXmlStyle(document);
			log.debug("2.完成解析XML中样式信息");
			List<Worksheet> worksheets = readXmlWorksheet(document);
			if (log.isDebugEnabled()) {
				log.debug("3.开始将XML信息写入Excel，数据为：" + worksheets.toString());
			}
			HSSFWorkbook wb = new HSSFWorkbook();
			for (Worksheet worksheet : worksheets) {
				HSSFSheet sheet = wb.createSheet(worksheet.getName());
				Table table = worksheet.getTable();
				List<Row> rows = table.getRows();
				List<Column> columns = table.getColumns();

				if (columns != null && columns.size() > 0) {
					// 填充列宽
					int columnIndex = 0;
					for (int i = 0; i < columns.size(); i++) {
						Column column = columns.get(i);
						columnIndex = getCellWidthIndex(columnIndex, i, column.getIndex());
						sheet.setColumnWidth(columnIndex, (int) column.getWidth() * 50);
					}
				}

				int createRowIndex = 0;
				List<CellRangeAddressEntity> cellRangeAddresses = new ArrayList<>();
				for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
					Row rowInfo = rows.get(rowIndex);
					if (rowInfo == null) {
						continue;
					}
					createRowIndex = getIndex(createRowIndex, rowIndex, rowInfo.getIndex());
					HSSFRow row = sheet.createRow(createRowIndex);
					if (rowInfo.getHeight() != null) {
						Integer height = rowInfo.getHeight() * 20;
						row.setHeight(height.shortValue());
					}
					List<Cell> cells = rowInfo.getCells();
					if (CollectionUtils.isEmpty(cells)) {
						continue;
					}
					int startIndex = 0;
					for (int cellIndex = 0; cellIndex < cells.size(); cellIndex++) {
						Cell cellInfo = cells.get(cellIndex);
						if (cellInfo == null) {
							continue;
						}
						// 获取起始列
						startIndex = getIndex(startIndex, cellIndex, cellInfo.getIndex());
						HSSFCell cell = row.createCell(startIndex);
						String styleID = cellInfo.getStyleID();
						Style style = styleMap.get(styleID);
						/*设置数据单元格格式*/
						CellStyle dataStyle = wb.createCellStyle();
						// 设置边框样式
						setBorder(style, dataStyle);
						// 设置对齐方式
						setAlignment(style, dataStyle);
						// 填充文本
						setValue(wb, cellInfo, cell, style, dataStyle);
						// 填充颜色
						setCellColor(style, dataStyle);
						cell.setCellStyle(dataStyle);
						//单元格注释
						if (cellInfo.getComment() != null) {
							Data data = cellInfo.getComment().getData();
							Comment comment = sheet.createDrawingPatriarch()
									.createCellComment(new HSSFClientAnchor(0, 0, 0, 0, (short) 3, 3, (short) 5, 6));
							comment.setString(new HSSFRichTextString(data.getText()));
							cell.setCellComment(comment);
						}
						// 合并单元格
						startIndex = getCellRanges(createRowIndex, cellRangeAddresses, startIndex, cellInfo, style);
					}
				}
				// 添加合并单元格
				addCellRange(sheet, cellRangeAddresses);
				// 添加下拉
				addDataValidation(sheet,worksheet);
			}
			// 加载图片到excel
			log.debug("4.开始写入图片：" + freemarkerInput.getExcelImageInputs());
			if (!CollectionUtils.isEmpty(freemarkerInput.getExcelImageInputs())) {
				writeImageToExcel(freemarkerInput.getExcelImageInputs(), wb);
			}
			log.debug("5.完成写入图片：" + freemarkerInput.getExcelImageInputs());
			// 写入excel文件,response字符流转换成字节流，template需要字节流作为输出
			wb.write(outputStream);
			outputStream.close();
		} catch (Exception e) {
			e.printStackTrace();
			log.error("导出excel异常：" + e.getMessage());
		} finally {
			try {
				out.close();
			} catch (Exception e) {

			}
		}
	}

	/**
	 * 导出Excel到输出流（支持Excel2007版，xlsx格式）
	 *
	 * @param freemarkerInput
	 * @param outputStream
	 */
	public static void createExcelToStream(FreemarkerInput freemarkerInput, OutputStream outputStream) {
		Writer out = null;
		try {
			// 创建xml文件
			Template template = getTemplate(freemarkerInput.getTemplateName(), freemarkerInput.getTemplateFilePath());
			File tempXMLFile = new File(freemarkerInput.getXmlTempFile() + freemarkerInput.getFileName() + ".xml");
			FileUtils.forceMkdirParent(tempXMLFile);
			out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(tempXMLFile), "UTF-8"));
			template.process(freemarkerInput.getDataMap(), out);
			if (log.isDebugEnabled()) {
				log.debug("1.完成将文本数据导入到XML文件中");
			}
			SAXReader reader = new SAXReader();
			Document document = reader.read(tempXMLFile);
			Map<String, Style> styleMap = readXmlStyle(document);
			log.debug("2.完成解析XML中样式信息");
			List<Worksheet> worksheets = readXmlWorksheet(document);
			if (log.isDebugEnabled()) {
				log.debug("3.开始将XML信息写入Excel，数据为：" + worksheets.toString());
			}
			XSSFWorkbook wb = new XSSFWorkbook();
			for (Worksheet worksheet : worksheets) {
				XSSFSheet sheet = wb.createSheet(worksheet.getName());
				Table table = worksheet.getTable();
				List<Row> rows = table.getRows();
				List<Column> columns = table.getColumns();
				if (columns != null && columns.size() > 0) {
					// 填充列宽
					int columnIndex = 0;
					for (int i = 0; i < columns.size(); i++) {
						Column column = columns.get(i);
						columnIndex = getCellWidthIndex(columnIndex, i, column.getIndex());
						sheet.setColumnWidth(columnIndex, (int) column.getWidth() * 50);
					}
				}
				int createRowIndex = 0;
				List<CellRangeAddressEntity> cellRangeAddresses = new ArrayList<>();
				for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
					Row rowInfo = rows.get(rowIndex);
					if (rowInfo == null) {
						continue;
					}
					createRowIndex = getIndex(createRowIndex, rowIndex, rowInfo.getIndex());
					XSSFRow row = sheet.createRow(createRowIndex);
					if (rowInfo.getHeight() != null) {
						Integer height = rowInfo.getHeight() * 20;
						row.setHeight(height.shortValue());
					}
					List<Cell> cells = rowInfo.getCells();
					if (CollectionUtils.isEmpty(cells)) {
						continue;
					}
					int startIndex = 0;
					for (int cellIndex = 0; cellIndex < cells.size(); cellIndex++) {
						Cell cellInfo = cells.get(cellIndex);
						if (cellInfo == null) {
							continue;
						}
						// 获取起始列
						startIndex = getIndex(startIndex, cellIndex, cellInfo.getIndex());
						XSSFCell cell = row.createCell(startIndex);
						String styleID = cellInfo.getStyleID();
						Style style = styleMap.get(styleID);
						/*设置数据单元格格式*/
						CellStyle dataStyle = wb.createCellStyle();
						// 设置边框样式
						setBorder(style, dataStyle);
						// 设置对齐方式
						setAlignment(style, dataStyle);
						// 填充文本
						setValue(wb, cellInfo, cell, style, dataStyle);
						// 填充颜色
						setCellColor(style, dataStyle);
						cell.setCellStyle(dataStyle);
						//单元格注释
						if (cellInfo.getComment() != null) {
							Data data = cellInfo.getComment().getData();
							Comment comment = sheet.createDrawingPatriarch()
									.createCellComment(new XSSFClientAnchor(0, 0, 0, 0, (short) 3, 3, (short) 5, 6));
							comment.setString(new XSSFRichTextString(data.getText()));
							cell.setCellComment(comment);
						}
						// 合并单元格
						startIndex = getCellRanges(createRowIndex, cellRangeAddresses, startIndex, cellInfo, style);
					}
				}
				// 添加合并单元格
				addCellRange(sheet, cellRangeAddresses);
				// 添加下拉
				addDataValidation(sheet,worksheet);

			}
			List<String> hideSheetNames = freemarkerInput.getHideSheetNames();
			//设置隐藏
			if(hideSheetNames!=null && !hideSheetNames.isEmpty() ){
				for (String hideSheetName : hideSheetNames) {
					int sheetIndex = wb.getSheetIndex(hideSheetName);
					wb.setSheetHidden(sheetIndex,true);
				}
			}
			// 加载图片到excel
			log.debug("4.开始写入图片：" + freemarkerInput.getExcelImageInputs());
			if (!CollectionUtils.isEmpty(freemarkerInput.getExcelImageInputs())) {
				writeImageToExcel(freemarkerInput.getExcelImageInputs(), wb);
			}
			log.debug("5.完成写入图片：" + freemarkerInput.getExcelImageInputs());
			// 写入excel文件,response字符流转换成字节流，template需要字节流作为输出
			wb.write(outputStream);
			outputStream.close();
		} catch (Exception e) {
			e.printStackTrace();
			log.error("导出excel异常：" + e.getMessage());
		} finally {
			try {
				out.close();
			} catch (Exception e) {
			}
		}
	}

	private static void addDataValidation(Sheet sheet, Worksheet worksheet) {
		List<DataValidations> dataValidationList = worksheet.getDataValidationList();
		if(dataValidationList!=null && !dataValidationList.isEmpty()){
			for (DataValidations dataValidations : dataValidationList) {
				CellRangeAddressList rangeAddressList =
						new CellRangeAddressList(dataValidations.getStartRow()-1,
								dataValidations.getEndRow()-1,dataValidations.getStartCol()-1,dataValidations.getEndCol()-1);
				DataValidationHelper dataValidationHelper = sheet.getDataValidationHelper();
				DataValidationConstraint dataValidationConstraint = dataValidations.getIsformula()?
						dataValidationHelper.createFormulaListConstraint(dataValidations.getValueText()):
						dataValidationHelper.createExplicitListConstraint(dataValidations.getValueText().split(","));
				DataValidation dataValidation = dataValidationHelper.createValidation(dataValidationConstraint, rangeAddressList);
				dataValidation.setSuppressDropDownArrow(true);
				dataValidation.setShowPromptBox(true);
				dataValidation.setEmptyCellAllowed(true);
				sheet.addValidationData(dataValidation);
			}
		}
	}

	public static Map<String, Style> readXmlStyle(Document document) {
		Map<String, Style> styleMap = XmlReader.getStyle(document);
		return styleMap;
	}

	public static List<Worksheet> readXmlWorksheet(Document document) {
		List<Worksheet> worksheets = XmlReader.getWorksheet(document);
		return worksheets;
	}

	private static int getIndex(int columnIndex, int i, Integer index) {
		if (index != null) {
			columnIndex = index - 1;
		}
		if (index == null && columnIndex != 0) {
			columnIndex = columnIndex + 1;
		}
		if (index == null && columnIndex == 0) {
			columnIndex = i;
		}
		return columnIndex;
	}

	private static int getCellWidthIndex(int columnIndex, int i, Integer index) {
		if (index != null) {
			columnIndex = index;
		}
		if (index == null && columnIndex != 0) {
			columnIndex = columnIndex + 1;
		}
		if (index == null && columnIndex == 0) {
			columnIndex = i;
		}
		return columnIndex;
	}

	/**
	 * 设置边框
	 *
	 * @param style:
	 * @param dataStyle:
	 * @return void
	 */
	private static void setBorder(Style style, CellStyle dataStyle) {
		if (style != null && style.getBorders() != null) {
			for (int k = 0; k < style.getBorders().size(); k++) {
				Style.Border border = style.getBorders().get(k);
				if (border != null) {
					if ("Bottom".equals(border.getPosition())) {
						dataStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
						dataStyle.setBorderBottom(BorderStyle.THIN);
					}
					if ("Left".equals(border.getPosition())) {
						dataStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
						dataStyle.setBorderLeft(BorderStyle.THIN);
					}
					if ("Right".equals(border.getPosition())) {
						dataStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
						dataStyle.setBorderRight(BorderStyle.THIN);
					}
					if ("Top".equals(border.getPosition())) {
						dataStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
						dataStyle.setBorderTop(BorderStyle.THIN);
					}
				}

			}
		}
	}

	/**
	 * 将图片写入Excel(XLS版)
	 *
	 * @param excelImageInputs
	 * @param wb
	 * @throws IOException
	 */
	@SuppressWarnings("rawtypes")
	private static void writeImageToExcel(List<ExcelImageInput> excelImageInputs, HSSFWorkbook wb) throws IOException {
		BufferedImage bufferImg = null;
		if (!CollectionUtils.isEmpty(excelImageInputs)) {
			for (ExcelImageInput excelImageInput : excelImageInputs) {
				Sheet sheet = wb.getSheetAt(excelImageInput.getSheetIndex());
				if (sheet == null) {
					continue;
				}
				// 画图的顶级管理器，一个sheet只能获取一个
				Drawing patriarch = sheet.createDrawingPatriarch();
				// anchor存储图片的属性，包括在Excel中的位置、大小等信息
				HSSFClientAnchor anchor = excelImageInput.getAnchorXls();
				anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
				// 插入图片
				String imagePath = excelImageInput.getImgPath();
				// 将图片写入到byteArray中
				ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
				bufferImg = ImageIO.read(new File(imagePath));
				// 图片扩展名
				String imageType = imagePath.substring(imagePath.lastIndexOf(".") + 1, imagePath.length());
				ImageIO.write(bufferImg, imageType, byteArrayOut);
				// 通过poi将图片写入到Excel中
				patriarch.createPicture(anchor,
						wb.addPicture(byteArrayOut.toByteArray(), HSSFWorkbook.PICTURE_TYPE_JPEG));
			}
		}
	}

	/**
	 * 将图片写入Excel(XLSX版)
	 *
	 * @param excelImageInputs
	 * @param wb
	 * @throws IOException
	 */
	@SuppressWarnings("rawtypes")
	private static void writeImageToExcel(List<ExcelImageInput> excelImageInputs, XSSFWorkbook wb) throws IOException {
		BufferedImage bufferImg = null;
		if (!CollectionUtils.isEmpty(excelImageInputs)) {
			for (ExcelImageInput excelImageInput : excelImageInputs) {
				Sheet sheet = wb.getSheetAt(excelImageInput.getSheetIndex());
				if (sheet == null) {
					continue;
				}
				// 画图的顶级管理器，一个sheet只能获取一个
				Drawing patriarch = sheet.createDrawingPatriarch();
				// anchor存储图片的属性，包括在Excel中的位置、大小等信息
				XSSFClientAnchor anchor = excelImageInput.getAnchorXlsx();
				anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
				// 插入图片
				String imagePath = excelImageInput.getImgPath();
				// 将图片写入到byteArray中
				ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
				bufferImg = ImageIO.read(new File(imagePath));
				// 图片扩展名
				String imageType = imagePath.substring(imagePath.lastIndexOf(".") + 1, imagePath.length());
				ImageIO.write(bufferImg, imageType, byteArrayOut);
				// 通过poi将图片写入到Excel中
				patriarch.createPicture(anchor,
						wb.addPicture(byteArrayOut.toByteArray(), HSSFWorkbook.PICTURE_TYPE_JPEG));
			}
		}
	}

	/**
	 * 添加合并单元格（XLS格式）
	 *
	 * @param sheet:
	 * @param cellRangeAddresses:
	 * @return void
	 */
	private static void addCellRange(HSSFSheet sheet, List<CellRangeAddressEntity> cellRangeAddresses) {
		if (!CollectionUtils.isEmpty(cellRangeAddresses)) {
			for (CellRangeAddressEntity cellRangeAddressEntity : cellRangeAddresses) {
				CellRangeAddress cellRangeAddress = cellRangeAddressEntity.getCellRangeAddress();
				sheet.addMergedRegion(cellRangeAddress);
				if (CollectionUtils.isEmpty(cellRangeAddressEntity.getBorders())) {
					continue;
				}
				for (int k = 0; k < cellRangeAddressEntity.getBorders().size(); k++) {
					Style.Border border = cellRangeAddressEntity.getBorders().get(k);
					if (border == null) {
						continue;
					}
					if ("Bottom".equals(border.getPosition())) {
						RegionUtil.setBorderBottom(BorderStyle.THIN, cellRangeAddress, sheet);
					}
					if ("Left".equals(border.getPosition())) {
						RegionUtil.setBorderLeft(BorderStyle.THIN, cellRangeAddress, sheet);
					}
					if ("Right".equals(border.getPosition())) {
						RegionUtil.setBorderRight(BorderStyle.THIN, cellRangeAddress, sheet);
					}
					if ("Top".equals(border.getPosition())) {
						RegionUtil.setBorderTop(BorderStyle.THIN, cellRangeAddress, sheet);
					}
				}
			}
		}
	}

	/**
	 * 添加合并单元格（XLSX格式）
	 *
	 * @param sheet:
	 * @param cellRangeAddresses:
	 * @return void
	 */
	private static void addCellRange(XSSFSheet sheet, List<CellRangeAddressEntity> cellRangeAddresses) {
		if (!CollectionUtils.isEmpty(cellRangeAddresses)) {
			for (CellRangeAddressEntity cellRangeAddressEntity : cellRangeAddresses) {
				CellRangeAddress cellRangeAddress = cellRangeAddressEntity.getCellRangeAddress();
				sheet.addMergedRegion(cellRangeAddress);
				if (CollectionUtils.isEmpty(cellRangeAddressEntity.getBorders())) {
					continue;
				}
				for (int k = 0; k < cellRangeAddressEntity.getBorders().size(); k++) {
					Style.Border border = cellRangeAddressEntity.getBorders().get(k);
					if (border == null) {
						continue;
					}
					if ("Bottom".equals(border.getPosition())) {
						RegionUtil.setBorderBottom(BorderStyle.THIN, cellRangeAddress, sheet);
					}
					if ("Left".equals(border.getPosition())) {
						RegionUtil.setBorderLeft(BorderStyle.THIN, cellRangeAddress, sheet);
					}
					if ("Right".equals(border.getPosition())) {
						RegionUtil.setBorderRight(BorderStyle.THIN, cellRangeAddress, sheet);
					}
					if ("Top".equals(border.getPosition())) {
						RegionUtil.setBorderTop(BorderStyle.THIN, cellRangeAddress, sheet);
					}
				}
			}
		}
	}

	/**
	 * 设置对齐方式
	 *
	 * @param style:
	 * @param dataStyle:
	 * @return void
	 */
	private static void setAlignment(Style style, CellStyle dataStyle) {
		if (style != null && style.getAlignment() != null) {
			// 设置水平对齐方式
			String horizontal = style.getAlignment().getHorizontal();
			if (!ObjectUtils.isEmpty(horizontal)) {
				if ("Left".equals(horizontal)) {
					dataStyle.setAlignment(HorizontalAlignment.LEFT);
				} else if ("Center".equals(horizontal)) {
					dataStyle.setAlignment(HorizontalAlignment.CENTER);
				} else {
					dataStyle.setAlignment(HorizontalAlignment.RIGHT);
				}
			}

			// 设置垂直对齐方式
			String vertical = style.getAlignment().getVertical();
			if (!ObjectUtils.isEmpty(vertical)) {
				if ("Top".equals(vertical)) {
					dataStyle.setVerticalAlignment(VerticalAlignment.TOP);
				} else if ("Center".equals(vertical)) {
					dataStyle.setVerticalAlignment(VerticalAlignment.CENTER);
				} else if ("Bottom".equals(vertical)) {
					dataStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
				} else if ("JUSTIFY".equals(vertical)) {
					dataStyle.setVerticalAlignment(VerticalAlignment.JUSTIFY);
				} else {
					dataStyle.setVerticalAlignment(VerticalAlignment.DISTRIBUTED);
				}
			}
			// 设置换行
			String wrapText = style.getAlignment().getWrapText();
			if (!ObjectUtils.isEmpty(wrapText)) {
				dataStyle.setWrapText(true);
			}
		}
	}

	/**
	 * 设置单元格背景填充色
	 *
	 * @param style:
	 * @param dataStyle:
	 * @return void
	 */
	private static void setCellColor(Style style, CellStyle dataStyle) {
		if (style != null && style.getInterior() != null) {

			String color = style.getInterior().getColor();

			if (color == null) {
				color = "#FFFFFF";
			}

			Integer[] rgb = ColorUtil.hex2Rgb(color);

			HSSFWorkbook hssfWorkbook = new HSSFWorkbook();

			HSSFPalette palette = hssfWorkbook.getCustomPalette();

			HSSFColor paletteColor = palette.findSimilarColor(rgb[0],rgb[1],rgb[2]);

			dataStyle.setFillForegroundColor(paletteColor.getIndex());

			dataStyle.setFillBackgroundColor(paletteColor.getIndex());

			if ("Solid".equals(style.getInterior().getPattern())) {
				dataStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			}
		}
	}

	/**
	 * 构造合并单元格集合
	 *
	 * @param createRowIndex:
	 * @param cellRangeAddresses:
	 * @param startIndex:
	 * @param cellInfo:
	 * @param style:
	 * @return int
	 */
	private static int getCellRanges(int createRowIndex, List<CellRangeAddressEntity> cellRangeAddresses,
			int startIndex, Cell cellInfo, Style style) {
		if (cellInfo.getMergeAcross() != null || cellInfo.getMergeDown() != null) {
			CellRangeAddress cellRangeAddress = null;
			if (cellInfo.getMergeAcross() != null && cellInfo.getMergeDown() != null) {
				int mergeAcross = startIndex;
				if (cellInfo.getMergeAcross() != 0) {
					// 获取该单元格结束列数
					mergeAcross += cellInfo.getMergeAcross();
				}
				int mergeDown = createRowIndex;
				if (cellInfo.getMergeDown() != 0) {
					// 获取该单元格结束列数
					mergeDown += cellInfo.getMergeDown();
				}
				cellRangeAddress = new CellRangeAddress(createRowIndex, mergeDown, (short) startIndex,
						(short) mergeAcross);
			} else if (cellInfo.getMergeAcross() != null && cellInfo.getMergeDown() == null) {
				int mergeAcross = startIndex;
				if (cellInfo.getMergeAcross() != 0) {
					// 获取该单元格结束列数
					mergeAcross += cellInfo.getMergeAcross();
					// 合并单元格
					cellRangeAddress = new CellRangeAddress(createRowIndex, createRowIndex, (short) startIndex,
							(short) mergeAcross);
				}

			} else if (cellInfo.getMergeDown() != null && cellInfo.getMergeAcross() == null) {
				int mergeDown = createRowIndex;
				if (cellInfo.getMergeDown() != 0) {
					// 获取该单元格结束列数
					mergeDown += cellInfo.getMergeDown();
					// 合并单元格
					cellRangeAddress = new CellRangeAddress(createRowIndex, mergeDown, (short) startIndex,
							(short) startIndex);
				}
			}

			if (cellInfo.getMergeAcross() != null) {
				int length = cellInfo.getMergeAcross().intValue();
				for (int i = 0; i < length; i++) {
					//startIndex += cellInfo.getMergeAcross();
					startIndex += 1;
				}
			}
			CellRangeAddressEntity cellRangeAddressEntity = new CellRangeAddressEntity();
			cellRangeAddressEntity.setCellRangeAddress(cellRangeAddress);
			if (style != null && style.getBorders() != null) {
				cellRangeAddressEntity.setBorders(style.getBorders());
			}
			cellRangeAddresses.add(cellRangeAddressEntity);
		}
		return startIndex;
	}

	/**
	 * 设置文本值内容（XLSX格式）
	 *
	 * @param wb:
	 * @param cellInfo:
	 * @param cell:
	 * @param style:
	 * @param dataStyle:
	 * @return void
	 */
	private static void setValue(XSSFWorkbook wb, Cell cellInfo, XSSFCell cell, Style style, CellStyle dataStyle) {
		if (cellInfo.getData() != null) {
			XSSFFont font = wb.createFont();
			if (style != null && style.getFont() != null) {

				String color = style.getFont().getColor();

				if (color == null) {
					color = "#000000";
				}

				Integer[] rgb = ColorUtil.hex2Rgb(color);

				HSSFWorkbook hssfWorkbook = new HSSFWorkbook();

				HSSFPalette palette = hssfWorkbook.getCustomPalette();

				HSSFColor paletteColor = palette.findSimilarColor(rgb[0],rgb[1],rgb[2]);

				font.setColor(paletteColor.getIndex());
			}
			if (!ObjectUtils.isEmpty(cellInfo.getData().getType()) && "Number".equals(cellInfo.getData().getType())) {
				cell.setCellType(CellType.NUMERIC);
			}
			if (style != null && style.getFont().getBold() > 0) {
				font.setBold(true);
			}
			if (style != null && !ObjectUtils.isEmpty(style.getFont().getFontName())) {
				font.setFontName(style.getFont().getFontName());
			}
			if (style != null && style.getFont().getSize() > 0) {
				// 设置字体大小道
				font.setFontHeightInPoints((short) style.getFont().getSize());
			}

			if (cellInfo.getData().getFont() != null) {
				if (cellInfo.getData().getFont().getBold() > 0) {
					font.setBold(true);
				}
				if ("Number".equals(cellInfo.getData().getType())) {
					cell.setCellValue(Float.parseFloat(cellInfo.getData().getFont().getText()));
				} else {
					cell.setCellValue(cellInfo.getData().getFont().getText());
				}
				if (!ObjectUtils.isEmpty(cellInfo.getData().getFont().getCharSet())) {
					font.setCharSet(Integer.valueOf(cellInfo.getData().getFont().getCharSet()));
				}
			} else {
				if ("Number".equals(cellInfo.getData().getType())) {
					if (!ObjectUtils.isEmpty(cellInfo.getData().getText())) {
						// cell.setCellValue(Float.parseFloat(cellInfo.getData().getText()));
						cell.setCellValue(Float.parseFloat(cellInfo.getData().getText().replaceAll(",", "")));
					}
				} else {
					cell.setCellValue(cellInfo.getData().getText());

				}
			}

			if (style != null) {
				if (style.getNumberFormat() != null) {

					String color = style.getFont().getColor();

					if (color == null) {
						color = "#000000";
					}

					Integer[] rgb = ColorUtil.hex2Rgb(color);

					HSSFWorkbook hssfWorkbook = new HSSFWorkbook();

					HSSFPalette palette = hssfWorkbook.getCustomPalette();

					HSSFColor paletteColor = palette.findSimilarColor(rgb[0],rgb[1],rgb[2]);

					font.setColor(paletteColor.getIndex());

					if ("0%".equals(style.getNumberFormat().getFormat())) {
						XSSFDataFormat format = wb.createDataFormat();
						dataStyle.setDataFormat(format.getFormat(style.getNumberFormat().getFormat()));
					} else {
						dataStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("#,##0.00"));
					}
				}
			}
			dataStyle.setFont(font);
		}
	}

	/**
	 * 设置文本值内容（XLS格式）
	 *
	 * @param wb:
	 * @param cellInfo:
	 * @param cell:
	 * @param style:
	 * @param dataStyle:
	 * @return void
	 */
	private static void setValue(HSSFWorkbook wb, Cell cellInfo, HSSFCell cell, Style style, CellStyle dataStyle) {
		if (cellInfo.getData() != null) {
			HSSFFont font = wb.createFont();
			if (style != null && style.getFont() != null) {
				String color = style.getFont().getColor();

				if (color == null) {
					color = "#000000";
				}

				Integer[] rgb = ColorUtil.hex2Rgb(color);

				HSSFWorkbook hssfWorkbook = new HSSFWorkbook();

				HSSFPalette palette = hssfWorkbook.getCustomPalette();

				HSSFColor paletteColor = palette.findSimilarColor(rgb[0],rgb[1],rgb[2]);

				font.setColor(paletteColor.getIndex());
			}
			if (!ObjectUtils.isEmpty(cellInfo.getData().getType()) && "Number".equals(cellInfo.getData().getType())) {
				cell.setCellType(CellType.NUMERIC);
			}
			if (style != null && style.getFont().getBold() > 0) {
				font.setBold(true);
			}
			if (style != null && !ObjectUtils.isEmpty(style.getFont().getFontName())) {
				font.setFontName(style.getFont().getFontName());
			}
			if (style != null && style.getFont().getSize() > 0) {
				// 设置字体大小道
				font.setFontHeightInPoints((short) style.getFont().getSize());
			}

			if (cellInfo.getData().getFont() != null) {
				if (cellInfo.getData().getFont().getBold() > 0) {
					font.setBold(true);
				}
				if ("Number".equals(cellInfo.getData().getType())) {
					cell.setCellValue(Float.parseFloat(cellInfo.getData().getFont().getText()));
				} else {
					cell.setCellValue(cellInfo.getData().getFont().getText());
				}
				if (!ObjectUtils.isEmpty(cellInfo.getData().getFont().getCharSet())) {
					font.setCharSet(Integer.valueOf(cellInfo.getData().getFont().getCharSet()));
				}
			} else {
				if ("Number".equals(cellInfo.getData().getType())) {
					if (!ObjectUtils.isEmpty(cellInfo.getData().getText())) {
						// cell.setCellValue(Float.parseFloat(cellInfo.getData().getText()));
						cell.setCellValue(Float.parseFloat(cellInfo.getData().getText().replaceAll(",", "")));
					}
				} else {
					cell.setCellValue(cellInfo.getData().getText());

				}
			}

			if (style != null) {
				if (style.getNumberFormat() != null) {
					String color = style.getFont().getColor();

					if (color == null) {
						color = "#000000";
					}

					Integer[] rgb = ColorUtil.hex2Rgb(color);

					HSSFWorkbook hssfWorkbook = new HSSFWorkbook();

					HSSFPalette palette = hssfWorkbook.getCustomPalette();

					HSSFColor paletteColor = palette.findSimilarColor(rgb[0],rgb[1],rgb[2]);

					font.setColor(paletteColor.getIndex());

					if ("0%".equals(style.getNumberFormat().getFormat())) {
						HSSFDataFormat format = wb.createDataFormat();
						dataStyle.setDataFormat(format.getFormat(style.getNumberFormat().getFormat()));
					} else {
						dataStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("#,##0.00"));
					}

				}
			}
			dataStyle.setFont(font);
		}
	}

	public static void main(String[] args) {
//		List<String> detpNameList = Stream.of("宏能微电子/销售部/深圳销售/Alex组,宏能微电子/销售部/深圳销售/Fiona组,宏能微电子/销售部/上海销售,宏能微电子/销售部/苏州销售/Rikky组,宏能微电子/销售部/苏州销售/Sara组,宏能微电子/销售部/武汉销售,宏能微电子/销售部/天津销售,宏能微电子/销售部/厦门销售,宏能微电子/销售部/南京销售,宏能微电子/市场部,宏能微电子/采购部/亚洲采购部,宏能微电子/物流部,宏能微电子/品质部,宏能微电子/行政部,壹探网络/研发部,壹探网络/数据部,壹探网络/运营部,宏能微电子/管理层,宏能微电子/财务部,宏能微电子/销售部/深圳销售/Summer组".split(",")).collect(Collectors.toList());
//		List<String> typeNameList = Stream.of("办公/电脑1,办公/电脑2,办公/电脑3,办公/电脑4,办公/电脑5".split(",")).collect(Collectors.toList());
//		String assetsTypeName = "资产分类";
//		String detpName = "使用部门";
//		Map<String,Object> data = new HashMap(){
//			{
//				put("admins", "deng-001,deng002,deng003");
//				put("detpName", detpName+"!$A$1:$A$"+detpNameList.size());
////				put("detpNames", "壹探网络/研发部,壹探网络/数据部,壹探网络/运营部,宏能微电子/管理层,宏能微电子/财务部,宏能微电子/销售部/深圳销售/Summer组,宏能微电子/销售部/深圳销售/Bella组,宏能微电子/销售部/深圳销售/Ella组,宏能微电子/销售部/深圳销售/Louis组");
//				put("deptNameList",detpNameList );
//				put("deptNameSize",detpNameList.size()+2);
//
//				put("companys", "公司1,公司2,公司3,公司4,公司5");
//				put("godownNames", "深圳,北京,武汉,大连,厦门");
//				put("assetsSources", "采购,租赁");
//				put("assetsTypeName",assetsTypeName+"!$A$1:$A$"+typeNameList.size());
//				put("assetsTypeNameList",typeNameList);
//				put("assetsTypeNameSize",typeNameList.size()+2);
//			}};
//////         这个导出没有下拉
//		FreemarkerInput freemarkerInput = new FreemarkerInput();
//		//设置freemarker模板路径
//		freemarkerInput.setTemplateFilePath("/");
//		freemarkerInput.setHideSheetNames(Stream.of(assetsTypeName,detpName).collect(Collectors.toList()));
//		//模板名字
//		freemarkerInput.setTemplateName("资产导入模块.ftl");
//		//缓存xml路径
//		freemarkerInput.setXmlTempFile(System.getProperty("java.io.tmpdir"));
//		//缓存xml名字
//		freemarkerInput.setFileName(System.currentTimeMillis() +"_tmpXml");
//		//设置freemarker模板数据
//		freemarkerInput.setDataMap(data);
//
//		//导出Excel到输出流（Excel2007版及以上，xlsx格式）速度快
//		FreemarkerUtils.exportImageExcelNew("D:\\code\\export-test\\资产导入-模板.xlsx", freemarkerInput);

	}
}
