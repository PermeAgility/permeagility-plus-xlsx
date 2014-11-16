package permeagility.plus.xlsx;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import permeagility.util.DatabaseConnection;
import permeagility.util.QueryResult;
import permeagility.web.Weblet;

import com.orientechnologies.orient.core.metadata.schema.OClass;
import com.orientechnologies.orient.core.metadata.schema.OSchema;
import com.orientechnologies.orient.core.metadata.schema.OType;
import com.orientechnologies.orient.core.record.impl.ODocument;
import com.orientechnologies.orient.core.record.impl.ORecordBytes;

public class ImportFile extends Weblet {

	public String getPage(DatabaseConnection con, HashMap<String, String> parms) {
		
		String toLoad = parms.get("LOAD");
		String sheetToLoad = parms.get("SHEET");
		String rowToLoad = parms.get("ROW");  // Defines the row to use as column headers
		String tableName = parms.get("TABLENAME");
		String go = parms.get("GO");
		if (toLoad != null) {
			System.out.println("toLoad="+toLoad);
			ODocument d = con.get(toLoad);
			if (d != null) {
				StringBuffer contentType = new StringBuffer();
				StringBuffer contentFilename = new StringBuffer();
				byte[] data = getFile(d,"file",contentFilename, contentType);
				System.out.println("filename="+contentFilename+" type="+contentType+" size="+data.length);
				if (contentFilename.toString().toLowerCase().endsWith("xlsx")) {
					try {
						Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(data));
						if (sheetToLoad == null) {
							StringBuffer sb = new StringBuffer();
					    	try {
					    		sb.append(tableStart(0)+row(tableHead("Sheet")+tableHead("Rows")+tableHead("Columns")+tableHead("Load")));
								for (int k = 0; k < wb.getNumberOfSheets(); k++) {
									Sheet sheet = wb.getSheetAt(k);
									System.out.println("sheet="+sheet.getSheetName());
						    		sb.append(row("data",column(sheet.getSheetName())
						    				+column(""+sheet.getPhysicalNumberOfRows())
						    				+column(""+sheet.getRow(0).getPhysicalNumberOfCells())
						    				+column(form(hidden("LOAD",toLoad)+input("TABLENAME",makePrettyCamelCase(sheet.getSheetName()))+button("SHEET",""+k,"Load")))));
								}
						    	sb.append(tableEnd());
					    	} catch (Exception e) {  
					    		e.printStackTrace();
					    		sb.append("Error opening spreadsheet: "+e.getMessage());
					    	}
							parms.put("SERVICE", "Select sheet to import");
							return 	head("Import")+body(standardLayout(con, parms, sb.toString()));
							
						} else {
							if (rowToLoad == null) {
								System.out.println("Loading sheet "+sheetToLoad);
								StringBuilder sb = new StringBuilder();
								try {
									int sht = Integer.parseInt(sheetToLoad);
									Sheet sheet = wb.getSheetAt(sht);
						    		sb.append(tableStart(0));
						    		int maxCell = 0;
									for (int r = 0; r < sheet.getPhysicalNumberOfRows(); r++) {
										Row row = sheet.getRow(r);
										int cc = (row == null ? 0 : row.getPhysicalNumberOfCells());
										if (cc > maxCell) maxCell = cc;
									}
									StringBuilder sbh = new StringBuilder();
									sbh.append(column(""));
									for (int x = 0; x < maxCell; x++) {
										Character colChar = new Character((char)('A'+x));
										sbh.append(tableHead(""+colChar));
									}	
									sb.append(row(sbh.toString()));
						    		for (int r = 0; r < sheet.getPhysicalNumberOfRows(); r++) {
										Row row = sheet.getRow(r);
										if (row != null) {
											StringBuilder sbr = new StringBuilder();
											sbr.append(tableHead(""+r));
											for (int c = 0; c < row.getPhysicalNumberOfCells();c++) {
												Cell cell = row.getCell(c);
												if (cell != null) {
													String cellValue = cell.toString();
													if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
														int t = cell.getCachedFormulaResultType();
														if (t == Cell.CELL_TYPE_NUMERIC) {
															cellValue = ""+cell.getNumericCellValue();
														}
													}
													if (cellValue.endsWith(".0")) { cellValue = cellValue.substring(0, cellValue.length()-2); }
													sbr.append(column(cellValue));
												} else {
													sbr.append(column("null"));
												}
											}
											sb.append(rowOnClick("clickable", sbr.toString(), "window.location.href='" + this.getClass().getName()
													+ "?LOAD="+toLoad+"&SHEET=" + sheetToLoad + "&TABLENAME=" + parms.get("TABLENAME") + "&ROW=" + r + "';"));
										}
									}
							    	sb.append(tableEnd());
								} catch (Exception e) {
									e.printStackTrace();
								}
								parms.put("SERVICE", "Select row with column definitions");
								return head("ImportSheet")+body(standardLayout(con, parms, sb.toString()));
								
							} else {
								StringBuilder sb = new StringBuilder();
								sb.append("This is what I would create<BR>");
								boolean create = false;
								if (go != null && go.equals("YES")) {
									System.out.println("****** Creating table for realz now ********");
									create = true;
								}
								try {
									int sht = Integer.parseInt(sheetToLoad);
									Sheet sheet = wb.getSheetAt(sht);
									String sheetName = sheet.getSheetName();
									sb.append("table="+sheetName+" and I would call it "+tableName+"<BR>");
									OSchema schema = null;
									if (create) schema = con.getDb().getMetadata().getSchema();
									if (schema != null && schema.existsClass(tableName)) {
										return "Table name already exists - go back and change it";
									}
									OClass cls = null;
									if (create && schema != null) cls = schema.createClass(tableName);
									int rw = Integer.parseInt(rowToLoad);
									Row row = sheet.getRow(rw);
									String colNames[] = new String[row.getPhysicalNumberOfCells()];
									for (int c = 0; c < row.getPhysicalNumberOfCells(); c++) {
										Cell cell = row.getCell(c);
										String colName = cell.toString();
										colNames[c] = makePrettyCamelCase(colName);
										if (colName != null && !colName.equals("")) {
											sb.append("column="+colName+" and call it "+makePrettyCamelCase(colName)+"<BR>");
											if (cls != null) cls.createProperty(makePrettyCamelCase(colName), OType.STRING);
										}
									}
									if (schema != null) {
										schema.save();
									}
									int rowCount = 0;
									int insertedRowCount = 0;
									for (int dr = rw+1; dr<sheet.getPhysicalNumberOfRows(); dr++) {
										Row drow = sheet.getRow(dr);
										ODocument newDoc = null;
										if (create && cls != null) {
											newDoc = con.create(tableName);
										}
										for (int c = 0; c<drow.getPhysicalNumberOfCells(); c++) {
											Cell cell = drow.getCell(c);
											if (cell != null) {
												String cellValue = cell.toString();
												if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
													int t = cell.getCachedFormulaResultType();
													if (t == Cell.CELL_TYPE_NUMERIC) {
														cellValue = ""+cell.getNumericCellValue();
													}
												}
												if (cellValue.endsWith(".0")) { cellValue = cellValue.substring(0, cellValue.length()-2); }
												if (create && newDoc != null) {
													newDoc.field(colNames[c],cellValue);
												}
											} else {
											}											
										}
										if (create && newDoc != null) {
											newDoc.save();
											insertedRowCount++;
										}
										rowCount++;
									}
									sb.append("<BR>"+rowCount + " rows ("+insertedRowCount+" inserted)");
								} catch (Exception e) {
									e.printStackTrace();
									return e.getStackTrace().toString();
								}
								sb.append("<BR>");
			    				if (!create) {
			    					sb.append(form(button("GO","YES","Create table")));
			    				} else {
			    					sb.append(link("permeagility.web.Table?TABLENAME="
											+tableName ,"Go to table"));
			    				}
								parms.put("SERVICE", "Select confirm to create the table");
								return head("ImportSheet")+body(standardLayout(con, parms, sb.toString()));

							}
						}
					} catch (Exception e) {
						e.printStackTrace();
					}
				} else {
					return head("ImportSheet")+body(standardLayout(con, parms, paragraph("Please use an xlsx file")));					
				}
			}
		}
		
    	StringBuffer sb = new StringBuffer();
    	try {
    		sb.append(tableStart(0)+row(tableHead("Name")+tableHead("Load")));
	    	QueryResult qr = con.query("SELECT FROM importedFiles");
	    	for (int i=0; i<qr.size(); i++) {
	    		sb.append(row(column(qr.getStringValue(i, "name"))+column(form(button("LOAD",qr.get(i).getIdentity().toString().substring(1),"Load")))));
	    	}
	    	sb.append(tableEnd());
    	} catch (Exception e) {  
    		e.printStackTrace();
    		sb.append("Error retrieving import files: "+e.getMessage());
    	}
		parms.put("SERVICE", "Select import file");
		return 	head("Context")+body(standardLayout(con, parms, sb.toString()));
	}

	
	public byte[] getFile(ODocument d, String column, StringBuffer contentFilename, StringBuffer contentType) {
		ORecordBytes bytes = d.field("file");
		if (bytes != null) {
			try {
				ByteArrayInputStream bis = new ByteArrayInputStream(bytes.toStream());
				StringBuffer content_type = new StringBuffer();
				if (bis.available() > 0) {
					int binc = bis.read();
					do {
						content_type.append((char)binc);
						binc = bis.read();
					} while (binc != 0x00 && bis.available() > 0);
				}
				StringBuffer content_filename = new StringBuffer();
				if (bis.available() > 0) {
					int binc = bis.read();
					do {
						content_filename.append((char)binc);
						binc = bis.read();
					} while (binc != 0x00 && bis.available() > 0);
				}
				contentFilename.append(content_filename);
				contentType.append(content_type);
				ByteArrayOutputStream content = new ByteArrayOutputStream();
				System.out.print("Reading blob content: available="+bis.available());
				int avail;
				int binc;
				while ((binc = bis.read()) != -1 && (avail = bis.available()) > 0) {
					content.write(binc);
					byte[] buf = new byte[avail];
					int br = bis.read(buf);
					if (br > 0) {
						content.write(buf,0,br);
					}
				}
				if (content.size() > 0) {
					System.out.println("contentSize="+content.size());
					return content.toByteArray();
				}
			} catch (Exception e) {
				System.out.println("Excepting getting file image: "+e.getMessage());
				e.printStackTrace();
			}
		} else {
			System.out.println("Import: File is empty");
		}
		return null;
	}
	
	
	public static void main(String[] args) {
		if (args.length < 1) {
			System.err.println("At least one argument expected");
			return;
		}

		String fileName = args[0];
		try {
			if (args.length == 1) {
				Workbook wb = new XSSFWorkbook(new FileInputStream(fileName));

				System.out.println("Data dump:\n");

				for (int k = 0; k < wb.getNumberOfSheets(); k++) {
					Sheet sheet = wb.getSheetAt(k);
					int rows = sheet.getPhysicalNumberOfRows();
					System.out.println("Sheet " + k + " \"" + wb.getSheetName(k) + "\" has " + rows + " row(s).");
					for (int r = 0; r < rows; r++) {
						Row row = sheet.getRow(r);
						if (row == null) {
							continue;
						}

						int cells = row.getPhysicalNumberOfCells();
						System.out.println("\nROW " + row.getRowNum() + " has " + cells
								+ " cell(s).");
						for (int c = 0; c < cells; c++) {
							Cell cell = row.getCell(c);
							String value = null;
							if (cell != null) {
								switch (cell.getCellType()) {	
									case Cell.CELL_TYPE_FORMULA:
										value = "FORMULA value=" + cell.getCellFormula();
										break;
	
									case Cell.CELL_TYPE_NUMERIC:
										value = "NUMERIC value=" + cell.getNumericCellValue();
										break;
		
									case Cell.CELL_TYPE_STRING:
										value = "STRING value=" + cell.getStringCellValue();
										break;
	
									default:
								}
								System.out.println("CELL col=" + cell.getColumnIndex() + " VALUE=" + value);
							}
						}
					}
				}
			} else {
				System.out.println("Please specify a filename");
			}
			} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
