package wcec.excel_parser;

import java.util.ArrayList;
import java.util.List;

public class CellGroup {
	String groupName;
	Family leadFamily;
	Integer groupNumber;
	
	List<Family> memberFamily = new ArrayList<Family>();

	public String getGroupName() {
		return groupName;
	}

	public void setGroupName(String groupName) {
		this.groupName = groupName;
	}

	public List<Family> getMemberFamily() {
		return memberFamily;
	}

	public void setMemberFamily(List<Family> memberFamily) {
		this.memberFamily = memberFamily;
	}
	
	public void addMemberFamily(Family aFamily) {
		memberFamily.add(aFamily);
	}

	public Family getLeadFamily() {
		return leadFamily;
	}

	public void setLeadFamily(Family leadFamily) {
		this.leadFamily = leadFamily;
	}

	public Integer getGroupNumber() {
		return groupNumber;
	}

	public void setGroupNumber(Integer groupNumber) {
		this.groupNumber = groupNumber;
	}
	
	
	
	
}
