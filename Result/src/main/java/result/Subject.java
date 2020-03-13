package result;

import java.util.List;

public class Subject {

	private String name;

	private List<Integer> marks;

	private Integer subjectTotal;

	public Subject() {

	}

	public Subject(String name, List<Integer> marks) {
		super();
		this.name = name;
		this.marks = marks;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public List<Integer> getMarks() {
		return marks;
	}

	public void setMarks(List<Integer> marks) {
		this.marks = marks;
	}

	public Integer getSubjectTotal() {
		return subjectTotal;
	}

	public void setSubjectTotal(Integer subjectTotal) {
		this.subjectTotal = subjectTotal;
	}

	@Override
	public String toString() {
		return "Subject [name=" + name + ", marks=" + marks + ", subjectTotal="
				+ subjectTotal + "]";
	}

	

}
