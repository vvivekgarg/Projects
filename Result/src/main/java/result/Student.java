package result;

import java.util.ArrayList;
import java.util.List;

public class Student {

	private Integer rollNum;
	private String srn;

	private String name;

	private String fathersName;

	private List<Subject> subjects;

	public Student() {
	}

	public String getSrn() {
		return srn;
	}

	public void setSrn(String srn) {
		this.srn = srn;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public Integer getRollNum() {
		return rollNum;
	}

	public void setRollNum(Integer rollNum) {
		this.rollNum = rollNum;
	}

	public List<Subject> getSubjects() {
		return subjects;
	}

	public String getFathersName() {
		return fathersName;
	}

	public void setFathersName(String fathersName) {
		this.fathersName = fathersName;
	}

	public void addSubjects(Subject subject) {
		if (this.subjects == null)
			this.subjects = new ArrayList<Subject>();
		this.subjects.add(subject);
	}

	@Override
	public String toString() {
		return "Student [rollNum=" + rollNum + ", srn=" + srn + ", name="
				+ name + ", fathersName=" + fathersName
				+ ", subjects=" + subjects + "]";
	}

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((rollNum == null) ? 0 : rollNum.hashCode());
		return result;
	}

	@Override
	public boolean equals(Object obj) {
		if (this == obj)
			return true;
		if (obj == null)
			return false;
		if (getClass() != obj.getClass())
			return false;
		Student other = (Student) obj;
		if (rollNum == null) {
			if (other.rollNum != null)
				return false;
		} else if (!rollNum.equals(other.rollNum))
			return false;
		return true;
	}

}
