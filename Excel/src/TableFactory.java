
public class TableFactory {
	public static Table getTable(String operation) {
		switch(operation) {
			case "»Ø¿î":
				return new Table1();
		}
		return null;
	}
}
