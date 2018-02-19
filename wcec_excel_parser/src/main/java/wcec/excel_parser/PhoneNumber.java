package wcec.excel_parser;

public class PhoneNumber {
	String areaCode;
	String threeDigitPrefix;
	String lineNumber;
	public String getAreaCode() {
		return areaCode;
	}
	public void setAreaCode(String areaCode) {
		this.areaCode = areaCode;
	}
	public String getThreeDigitPrefix() {
		return threeDigitPrefix;
	}
	public void setThreeDigitPrefix(String threeDigitPrefix) {
		this.threeDigitPrefix = threeDigitPrefix;
	}
	public String getLineNumber() {
		return lineNumber;
	}
	public void setLineNumber(String lineNumber) {
		this.lineNumber = lineNumber;
	}
	
	public static PhoneNumber parsePhoneNumber(String aNumber) {
		PhoneNumber aPhoneNumber = new PhoneNumber(); 
		String []parts = aNumber.split("-");
		if (parts.length == 2) {
			aPhoneNumber.setAreaCode("302");
			aPhoneNumber.setThreeDigitPrefix(parts[0]);
			aPhoneNumber.setLineNumber(parts[1]);
		} else if (parts.length == 3) {
			aPhoneNumber.setAreaCode(parts[0]);
			aPhoneNumber.setThreeDigitPrefix(parts[1]);
			aPhoneNumber.setLineNumber(parts[2]);
		} else {
			throw new RuntimeException("Incorrectly formatted phone number: " + aNumber );
		}
		return aPhoneNumber;
	}

}
