package wcec.excel_parser;

public class Family { 

	Person person1;
	Person person2;
	boolean single;
	PhoneNumber phoneNumber;
	public Person getPerson1() {
		return person1;
	}
	public void setPerson1(Person person1) {
		this.person1 = person1;
	}
	public Person getPerson2() {
		return person2;
	}
	public void setPerson2(Person person2) {
		this.person2 = person2;
	}
	public boolean isSingle() {
		return single;
	}
	public void setSingle(boolean single) {
		this.single = single;
	}
	public PhoneNumber getPhoneNumber() {
		return phoneNumber;
	}
	public void setPhoneNumber(PhoneNumber phoneNumber) {
		this.phoneNumber = phoneNumber;
	}
	
	
}
