
public class TableFactory {
	public static Table getTable(String operation) {
		switch(operation) {
			case "�ؿ�":
				return new Table1();
		}
		return null;
	}
}
