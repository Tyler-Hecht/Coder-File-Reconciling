
public class DatavyuObject {
	
	private int ordinal;
	private int onset;
	private int offset;
	private String code;
	
	public DatavyuObject(int ordinal, int onset, int offset, String code) {
		this.ordinal = ordinal;
		this.onset = onset;
		this.offset = offset;
		this.code = code;
	}
	
	public int getOrdinal() {
		return this.ordinal;
	}
	public void setOrdinal(int o) {
		this.ordinal = o;
	}
	public int getOnset() {
		return this.onset;
	}
	public void setOnset(int o) {
		this.onset = o;
	}
	public int getOffset() {
		return this.offset;
	}
	public void setOffset(int o) {
		this.offset = o;
	}
	public String getCode() {
		return this.code;
	}
	public void setCode(String c) {
		this.code = c;
	}

}
