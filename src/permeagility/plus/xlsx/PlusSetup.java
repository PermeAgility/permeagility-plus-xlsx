package permeagility.plus.xlsx;

import java.util.HashMap;

import permeagility.util.DatabaseConnection;
import permeagility.util.Setup;

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
	public String getVersion() { return "0.2.3"; }
	
	public boolean isInstalled() { return INSTALLED; }
	
	public boolean install(DatabaseConnection con, HashMap<String,String> parms, StringBuilder errors) {
		OSchema oschema = con.getSchema();
		String newTableGroup = pickTableGroup(con, parms);
		
		if (isNullOrBlank(newTableGroup) || isNullOrBlank(parms.get("MENU")) || isNullOrBlank(parms.get("ROLES"))) {
			errors.append(paragraph("error","Please specify a table group, menu and the roles to access"));
			return false;
		}
		
		// Setup tables
		OClass table = Setup.checkCreateClass(con, oschema, TABLE, errors, newTableGroup);
		Setup.checkCreateProperty(con, table, "name", OType.STRING, errors);
		Setup.checkCreateProperty(con, table, "file", OType.STRING, errors);
		Setup.checkCreateProperty(con, table, "loaded", OType.DATETIME, errors);
		
		OClass logTable = Setup.checkCreateClass(con, oschema, LOGTABLE, errors, newTableGroup);
		Setup.checkCreateProperty(con, logTable, "file", OType.LINK, table, errors);
		Setup.checkCreateProperty(con, logTable, "sheet", OType.STRING, errors);
		Setup.checkCreateProperty(con, logTable, "row", OType.INTEGER, errors);
		Setup.checkCreateProperty(con, logTable, "className", OType.STRING, errors);
		Setup.checkCreateProperty(con, logTable, "created", OType.DATETIME, errors);
		Setup.checkCreateProperty(con, logTable, "executed", OType.DATETIME, errors);

		// Setup menu
		Setup.createMenuItem(con,getName(),getInfo(),MENU_CLASS,parms.get("MENU"),parms.get("ROLES"));	
		
		// Set the INSTALLED and INSTALLED_VERSION constant
		Setup.checkCreateConstant(con,this.getClass().getName(),getInfo(),"INSTALLED","true");
		Setup.checkCreateConstant(con,this.getClass().getName(),getInfo(),"INSTALLED_VERSION",getVersion());
		INSTALLED = true;
		return true;
	}
	
	public boolean remove(DatabaseConnection con, HashMap<String,String> parms, StringBuilder errors) {
		
		// Remove from menu
		if (parms.get("REMOVE_MENU") != null) {
			Setup.removeMenuItem(con, MENU_CLASS, errors);
		}
		
		// If specified, remove tables
		String remTab = parms.get("REMOVE_TABLES");
		if (remTab != null && remTab.equals("on")) {
			OSchema schema = con.getSchema();
			schema.dropClass(TABLE);
			errors.append(paragraph("error","Table dropped: "+TABLE));
			Setup.removeTableFromAllTableGroups(con, TABLE);
			schema.dropClass(LOGTABLE);
			errors.append(paragraph("error","Table dropped: "+LOGTABLE));
			Setup.removeTableFromAllTableGroups(con, LOGTABLE);
		}

		// Remove the INSTALLED and INSTALLED_VERSION constant
		Object ret2 = con.update("DELETE FROM "+Setup.TABLE_CONSTANT+" WHERE classname='"+this.getClass().getName()+"'");
		errors.append(paragraph("error","Delete INSTALLED and INSTALLED_VERSION constant "+ret2));
		INSTALLED = false;
		return true;
	}
	
	public boolean upgrade(DatabaseConnection con, HashMap<String,String> parms, StringBuilder errors) {
		// Perform upgrade actions
				
		// Then update the version constant
		Setup.checkCreateConstant(con,this.getClass().getName(),getInfo(),"INSTALLED_VERSION",getVersion());	
		return true;
	}

}
