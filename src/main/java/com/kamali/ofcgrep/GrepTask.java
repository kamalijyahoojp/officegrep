/**
 *
 */
package com.kamali.ofcgrep;

import java.io.File;
import java.io.FileFilter;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.lang.reflect.Field;
import java.text.MessageFormat;
import java.text.NumberFormat;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Pattern;

import javax.swing.SwingWorker;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFShapeGroup;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFShapeGroup;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;

/**
 * @author kamali
 *
 */
public class GrepTask extends SwingWorker<String, Object[]> {
	public interface ProcessNotify {
		public void process(List<Object[]> chunks);
	};
	private final MessageFormat SHAPE_COORD = new MessageFormat("({0,number},{1,number}){2}");
	private final Object[] coord = new Object[3];
	private final File dir;
	private final FileFilter excelFilter;
	private final Pattern strPattern;
	private final String shPattern;
	private final DataFormatter formatter = new DataFormatter();
	private int netTasks = 1;
	private int doneTasks = 0;
	private final Object[][] prg = new Object[8][3];
	private int prgIdx = 0;
	private ProcessNotify notify = null;
	/**
	 *
	 */
	public GrepTask(File dir, FileFilter excelFilter, String strPattern, String shPattern) {
		super();
		this.dir = dir;
		this.excelFilter = excelFilter;
		this.strPattern = Pattern.compile(strPattern);
		this.shPattern = shPattern;
	}

	public void setNotify(ProcessNotify notify) {
		this.notify = notify;
	}
	/* (非 Javadoc)
	 * @see javax.swing.SwingWorker#doInBackground()
	 */
	@Override
	protected String doInBackground() throws Exception {
		try {
			NumberFormat nf = (NumberFormat)SHAPE_COORD.getFormats()[0];
			nf.setGroupingUsed(false);
			nf = (NumberFormat)SHAPE_COORD.getFormats()[1];
			nf.setGroupingUsed(false);
			findDirectory(dir);
		} catch (Throwable e) {
			e.printStackTrace();
		}
		setProgress(100);
		return "done";
	}

	private void findDirectory(File dir) {
		File[] files = dir.listFiles(excelFilter);
		if (files == null) return;
		netTasks += files.length;
		for (File f : files) {
			if (f.isDirectory()) {
				--netTasks;
				findDirectory(f);
			} else {
				prg[prgIdx][0] = doneTasks;
				prg[prgIdx][1] = netTasks;
				prg[prgIdx][2] = f.getName();
				getPropertyChangeSupport().firePropertyChange("detail", null, prg[prgIdx]);
				prgIdx = (prgIdx + 1) % 8;
				try {
					excelSearch(f);
				} catch (Throwable e) {
					e.printStackTrace();
					Object[] info = new Object[] {
						f, e.getClass().getSimpleName(), "", e.getLocalizedMessage()
					};
					publish(info);
				}
				doneTasks++;
				setProgress(doneTasks * 100 / netTasks);
			}
		}
	}

	private void shapeSearch(File f, String sname, HSSFShape shape) {
		if (shape instanceof HSSFShapeGroup) {
			for (HSSFShape cshape : ((HSSFShapeGroup)shape).getChildren()) {
				shapeSearch(f, sname, cshape);
			}
		} else if (shape instanceof HSSFSimpleShape) {
			HSSFSimpleShape cshape = (HSSFSimpleShape)shape;
			try {
				HSSFRichTextString str = cshape.getString();
				if (str != null) {
					String text = str.getString();
					if (strPattern.matcher(text).find()) {
						String type = "?";
						switch (cshape.getShapeType()) {
						case HSSFSimpleShape.OBJECT_TYPE_ARC:
							type = "ARC";
							break;
						case HSSFSimpleShape.OBJECT_TYPE_COMBO_BOX:
							type = "ComboBox";
							break;
						case HSSFSimpleShape.OBJECT_TYPE_COMMENT:
							type = "COMMENT";
//							HSSFShape sh = cshape.getParent();
//							if (sh != null) {
//								shapeSearch(f, sname, sh);
//							}
							break;
						case HSSFSimpleShape.OBJECT_TYPE_LINE:
							type = "LINE";
							break;
						case HSSFSimpleShape.OBJECT_TYPE_MICROSOFT_OFFICE_DRAWING:
							type = "MS-Drawing";
							break;
						case HSSFSimpleShape.OBJECT_TYPE_OVAL:
							type = "OVAL";
							break;
						case HSSFSimpleShape.OBJECT_TYPE_PICTURE:
							type = "PICTURE";
							break;
						case HSSFSimpleShape.OBJECT_TYPE_RECTANGLE:
							type = "RECT";
							break;
						default:
							for (Field enums : HSSFSimpleShape.class.getDeclaredFields()) {
								enums.setAccessible(true);
								if (enums.getType() != int.class && enums.getType() != short.class)
									continue;
								int val = enums.getInt(cshape);
								if (val == cshape.getShapeType()) {
									type = enums.getName();
									break;
								}
							}
							break;
						}
						HSSFAnchor anchor = cshape.getAnchor();
						coord[0] = anchor.getDx1();
						coord[1] = anchor.getDy1();
						coord[2] = type;
						Object[] info = new Object[] {
							f, sname, SHAPE_COORD.format(coord), text.replace("\n", "<br>")
						};
						publish(info);
					}
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}

	private void shapeSearch(File f, String sname, XSSFShape shape) {
		if (shape instanceof XSSFShapeGroup) {
			XSSFDrawing draws = ((XSSFShapeGroup)shape).getDrawing();
			for (XSSFShape cshape : draws.getShapes()) {
				shapeSearch(f, sname, cshape);
			}
		} else if (shape instanceof XSSFSimpleShape) {
			XSSFSimpleShape cshape = (XSSFSimpleShape)shape;
			try {
				cshape.getAnchor();
				String text = cshape.getText();
				if (strPattern.matcher(text).find()) {
					String type = "?";
					for (Field enums : ShapeTypes.class.getDeclaredFields()) {
						enums.setAccessible(true);
						if (enums.getType() != int.class && enums.getType() != short.class)
							continue;
						int val = enums.getInt(cshape);
						if (val == cshape.getShapeType()) {
							type = enums.getName();
							break;
						}
					}
					XSSFAnchor anchor = cshape.getAnchor();
					coord[0] = anchor.getDx1() / 669;
					coord[1] = anchor.getDy1() / 669;
					coord[2] = type;
					Object[] info = new Object[] {
						f, sname, SHAPE_COORD.format(coord), text.replace("\n", "<br>")
					};
					publish(info);
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}

	private void excelSearch(File f) throws FileNotFoundException
		, EncryptedDocumentException, InvalidFormatException, IOException {
		String text = "";
		FileInputStream inp = new FileInputStream(f);

		Workbook wb = WorkbookFactory.create(inp);
		FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
		DataFormatter formatter = new DataFormatter();
//		HSSFWorkbook hssf = null;
//		SXSSFWorkbook sxssf = null;
//		XSSFWorkbook xssf = null;
//		if (wb instanceof HSSFWorkbook) {
//			hssf = (HSSFWorkbook)wb;
//			List<HSSFObjectData> objList = hssf.getAllEmbeddedObjects();
//		} else if (wb instanceof SXSSFWorkbook) {
//			sxssf = (SXSSFWorkbook)wb;
//		} else if (wb instanceof XSSFWorkbook) {
//			xssf = (XSSFWorkbook)wb;
//		}
		for (Iterator<Sheet> si = wb.sheetIterator();si.hasNext();) {
			Sheet sheet = si.next();
			String sname = sheet.getSheetName();
			Drawing draws = sheet.getDrawingPatriarch();
			if (shPattern != null && !sname.matches(shPattern))
				continue;
			for (Iterator<Row> ri = sheet.iterator();ri.hasNext();) {
				Row row = ri.next();
				for (Iterator<Cell> ci = row.iterator();ci.hasNext();) {
					Cell cell = ci.next();
					switch (cell.getCellType()) {
				    case FORMULA:
				    	text = cell.getCellFormula();
				    	try {
					    	text = formatter.formatCellValue(cell, evaluator);
				    	} catch (Throwable e) {
				        	System.out.println(e.toString());
				        	System.out.print("formula=");
				        	System.out.println(text);
				        }
				        break;
				    default:
				    	text = formatter.formatCellValue(cell);
				    	break;
					}
					if (strPattern.matcher(text).find()) {
						CellReference cellRef = new CellReference(cell);
						Object[] info = new Object[] {
							f, sname, cellRef.formatAsString(), text.replace("\n", "<br>")
						};
						publish(info);
					}
					Comment cm = cell.getCellComment();
					if (cm == null) continue;
					text = cm.getString().toString();
					if (strPattern.matcher(text).find()) {
						CellReference cellRef = new CellReference(cell);
						Object[] info = new Object[] {
							f, sname, cellRef.formatAsString() + "#comment", text.replace("\n", "<br>")
						};
						publish(info);
					}
				}
			}
			if (draws instanceof HSSFPatriarch) {
				HSSFPatriarch pa = (HSSFPatriarch)draws;
				for (HSSFShape shape : pa.getChildren()) {
					shapeSearch(f, sname, shape);
				}
			} else if (draws instanceof XSSFDrawing) {
				XSSFDrawing xdraw = (XSSFDrawing)draws;
				for (XSSFShape shape : xdraw.getShapes()) {
					shapeSearch(f, sname, shape);
				}
			}
		}
//		List<? extends PictureData> pics = wb.getAllPictures();
		wb.close();
		inp.close();
	}

	/* (非 Javadoc)
	 * @see javax.swing.SwingWorker#process(java.util.List)
	 */
	@Override
	protected void process(List<Object[]> chunks) {
		if (notify != null) notify.process(chunks);
	}

}
