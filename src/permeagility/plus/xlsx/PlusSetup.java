package permeagility.plus.xlsx;

import java.util.HashMap;

import permeagility.util.DatabaseConnection;
import permeagility.util.Setup;
import permeagility.web.Server;

import com.orientechnologies.orient.core.metadata.schema.OClass;
import com.orientechnologies.orient.core.metadata.schema.OSchema;
import com.orientechnologies.orient.core.metadata.schema.OType;

public class PlusSetup extends permeagility.plus.PlusSetup {

	public static boolean INSTALLED = false;  // Set via constant to complete installation
	public static String INSTALLED_VERSION = "0";  // Set via constant to complete installation
	
	// Override these to change the names of the tables that will be created and used by this importer
	public static String TABLE = "importedFiles";   // Local OrientDB table name to hold connection specs
	public static String LOGTABLE = "importedFilesLog";   // Saved path from a xlsx sheet    
	
	public static String MENU_CLASS = "permeagility.plus.xlsx.ImportFile";
	
	public String getName() { return "Import xlsx"; }
	public String getInfo() { return "Upload xlsx spreadsheet files and import sheets into tables"; }
	public String getVersion() { return "0.1.0"; }
	
	public boolean isInstalled() { return INSTALLED; }
	
	/** Install the plus module */
	public boolean install(DatabaseConnection con, HashMap<String,String> parms, StringBuilder errors) {
		OSchema oschema = con.getSchema();
		String newTableGroup = pickTableGroup(con, parms);
		
		if (isNullOrBlank(newTableGroup) || isNullOrBlank(parms.get("MENU")) || isNullOrBlank(parms.get("ROLES"))) {
			errors.append(paragraph("error","Please specify a table group, menu and the roles to access"));
			return false;
		}
		
		OClass table = Setup.checkCreateTable(con, oschema, TABLE, errors, newTableGroup);
		Setup.checkCreateColumn(con, table, "name", OType.STRING, errors);
		Setup.checkCreateColumn(con, table, "file", OType.CUSTOM, errors);
		Setup.checkCreateColumn(con, table, "loaded", OType.DATETIME, errors);
		
		OClass logTable = Setup.checkCreateTable(con, oschema, LOGTABLE, errors, newTableGroup);
		Setup.checkCreateColumn(con, logTable, "file", OType.LINK, table, errors);
		Setup.checkCreateColumn(con, logTable, "sheet", OType.STRING, errors);
		Setup.checkCreateColumn(con, logTable, "row", OType.INTEGER, errors);
		Setup.checkCreateColumn(con, logTable, "className", OType.STRING, errors);
		Setup.checkCreateColumn(con, logTable, "created", OType.DATETIME, errors);
		Setup.checkCreateColumn(con, logTable, "executed", OType.DATETIME, errors);

		Setup.createMenuItem(con,getName(),getInfo(),MENU_CLASS,parms.get("MENU"),parms.get("ROLES"));	
		
		setPlusInstalled(con, this.getClass().getName(), getInfo(), getVersion());
		INSTALLED = true;
		return true;
	}
	
	public boolean remove(DatabaseConnection con, HashMap<String,String> parms, StringBuilder errors) {
		
		if (parms.get("REMOVE_MENU") != null) {
			Setup.removeMenuItem(con, MENU_CLASS, errors);
		}

		String remTab = parms.get("REMOVE_TABLES");
		if (remTab != null && remTab.equals("on")) {
			Setup.dropTable(con, TABLE);
			errors.append(paragraph("success","Table dropped: "+TABLE));
			Setup.dropTable(con, LOGTABLE);
			errors.append(paragraph("success","Table dropped: "+LOGTABLE));
		}

		setPlusUninstalled(con, this.getClass().getName());
		INSTALLED = false;
		return true;
	}
	
	public boolean upgrade(DatabaseConnection con, HashMap<String,String> parms, StringBuilder errors) {
		// Perform upgrade actions

		
		setPlusVersion(con,this.getClass().getName(),getInfo(),getVersion());
		return true;
	}

}
