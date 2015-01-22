package permeagility.plus.xlsx;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import permeagility.util.DatabaseConnection;
import permeagility.util.Setup;
import permeagility.web.Message;
import permeagility.web.Server;
import permeagility.web.Table;

import com.orientechnologies.orient.core.metadata.schema.OClass;
import com.orientechnologies.orient.core.metadata.schema.OSchema;
import com.orientechnologies.orient.core.metadata.schema.OType;
import com.orientechnologies.orient.core.record.impl.ODocument;
import com.orientechnologies.orient.core.record.impl.ORecordBytes;

public class ImportFile extends Table {

	public String getPage(DatabaseConnection con, HashMap<String, String> parms) {
		
		StringBuilder errors = new StringBuilder();
		
		String submit = parms.get("SUBMIT");
		String editId = parms.get("EDIT_ID");
		String updateId = parms.get("UPDATE_ID");
		String toLoad = parms.get("LOAD");
		String sheetToLoad = parms.get("SHEET");
		String rowToLoad = parms.get("ROW");  // Defines the row to use as column headers
		String tableName = parms.get("TABLENAME");
		String go = parms.get("GO");
		
		// Show edit form if row selected for edit
		if (editId != null && tableName != null && submit == null && go == null) {
			return head("Edit", getDateControlScript(con.getLocale())+getColorControlScript())
					+ body(standardLayout(con, parms, getTableRowForm(con, tableName, parms)));
		}

		// Process insert - set loaded
		if (submit != null && submit.equals(Message.get(con.getLocale(), "CREATE_ROW"))) {
			parms.put(PARM_PREFIX+"loaded", formatDate(con.getLocale(), new java.util.Date(), "yyyy-MM-dd HH:mm:ss"));
			boolean inserted = insertRow(con,tableName,parms,errors);
			if (!inserted) {
				errors.append(paragraph("error","Could not insert"));
			}
		}		

		// Process update of work tables
		if (updateId != null && submit != null) {
			System.out.println("update_id="+updateId);
			if (submit.equals(Message.get(con.getLocale(), "DELETE"))) {
				if (deleteRow(con, tableName, parms, errors)) {
					submit = null;
				} else {
					return head("Could not delete")
							+ body(standardLayout(con, parms, getTableRowForm(con, tableName, parms) + errors.toString()));
				}
			} else if (submit.equals(Message.get(con.getLocale(), "UPDATE"))) {
				System.out.println("In updating row");
				if (updateRow(con, tableName, parms, errors)) {
				} else {
					return head("Could not update", getDateControlScript(con.getLocale())+getColorControlScript())
							+ body(standardLayout(con, parms, getTableRowForm(con, tableName, parms) + errors.toString()));
				}
			} 
			// Cancel is assumed
			editId = null;
			updateId = null;
		}

		if (toLoad != null) {
			System.out.println("toLoad="+toLoad);
			ODocument d = con.get(toLoad);
			if (d != null) {
				StringBuilder contentType = new StringBuilder();
				StringBuilder contentFilename = new StringBuilder();
				byte[] data = getFile(d,"file",contentFilename, contentType);
				System.out.println("filename="+contentFilename+" type="+contentType+" size="+data.length);
				if (contentFilename.toString().toLowerCase().endsWith("xlsx")) {
					try {
						Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(data));
						if (sheetToLoad == null) {
							StringBuilder sb = new StringBuilder();
					    	try {
					    		sb.append(tableStart(0)+row(columnHeader("Sheet")+columnHeader("Rows")+columnHeader("Columns")+columnHeader("Load")));
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
										sbh.append(columnHeader(""+colChar));
									}	
									sb.append(row(sbh.toString()));
						    		for (int r = 0; r < sheet.getPhysicalNumberOfRows(); r++) {
										Row row = sheet.getRow(r);
										if (row != null) {
											StringBuilder sbr = new StringBuilder();
											sbr.append(columnHeader(""+r));
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
											sb.append(rowOnClick("clickable", sbr.toString(), this.getClass().getName()
													+ "?LOAD="+toLoad+"&SHEET=" + sheetToLoad + "&TABLENAME=" + parms.get("TABLENAME") + "&ROW=" + r ));
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
											if (cls != null) Setup.checkCreateColumn(con, cls, makePrettyCamelCase(colName), OType.STRING, sb);
										}
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
		
    	StringBuilder sb = new StringBuilder();
		if (sb.length() == 0) {
	    	try {
	    		parms.put("SERVICE", "xlsxImporter: Setup/Select file");
				sb.append(paragraph("banner","Select Spreadsheet"));
				sb.append(getTable(con,parms,PlusSetup.TABLE,"SELECT FROM "+PlusSetup.TABLE, null,0, "name,loaded,button(LOAD:Load)"));
	    	} catch (Exception e) {  
	    		e.printStackTrace();
	    		sb.append("Error retrieving import files: "+e.getMessage());
	    	}
		}
		return 	head("Context",getDateControlScript(con.getLocale())+getColorControlScript())
				+body(standardLayout(con, parms, 
					errors.toString()
					+sb.toString()
					+((Server.getTablePriv(con, PlusSetup.TABLE) & PRIV_CREATE) > 0 && toLoad == null ? popupForm("CREATE_NEW_ROW",null,Message.get(con.getLocale(),"NEW_ROW"),null,"NAME",
							paragraph("banner",Message.get(con.getLocale(), "CREATE_ROW"))
							+hidden("TABLENAME", PlusSetup.TABLE)
							+getTableRowFields(con, PlusSetup.TABLE, parms, "name, file, -")
							+submitButton(Message.get(con.getLocale(), "CREATE_ROW"))) : "")
					));
	
	}

	
	public byte[] getFile(ODocument d, String column, StringBuilder contentFilename, StringBuilder contentType) {
		ORecordBytes bytes = d.field("file");
		if (bytes != null) {
			try {
				ByteArrayInputStream bis = new ByteArrayInputStream(bytes.toStream());
				StringBuilder content_type = new StringBuilder();
				if (bis.available() > 0) {
					int binc = bis.read();
					do {
						content_type.append((char)binc);
						binc = bis.read();
					} while (binc != 0x00 && bis.available() > 0);
				}
				StringBuilder content_filename = new StringBuilder();
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
	

}
