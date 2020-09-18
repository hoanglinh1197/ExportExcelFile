
public class InsertedQuery {
	private String schemaName;
	private String tableName;
	private String[] properties;
	private String[] value;

	public InsertedQuery(String schemaName, String tableName, String[] properties, String[] value) {
		super();
		this.schemaName = schemaName;
		this.properties = properties;
		this.value = value;
		this.tableName = tableName;
	}

	public String getQuery() {
		return "Insert into " + schemaName + "." + tableName + "(" + getCustomString(properties, true) + " values("
				+ getCustomString(value, false) + ";";
	}

	private String getCustomString(String[] strs, boolean isProperty) {
		StringBuilder str = new StringBuilder();
		for (String s : strs) {
			if (isProperty) {
				str.append(s + ",");
			} else {
				str.append(s == null ? s + "," : "'" + s + "',");

			}
		}
		str = new StringBuilder(str.substring(0, str.length() - 1));
		str.append(")");
		return str.toString();
	}

}
