package com.qorql.texttoexcel.response;

import java.util.List;

import com.qorql.texttoexcel.Patient;

public class Response {
	private List<Patient> patients;
	private List<Integer> rowNumbers;

	public List<Patient> getPatients() {
		return patients;
	}

	public void setPatients(List<Patient> patients) {
		this.patients = patients;
	}

	public List<Integer> getRowNumbers() {
		return rowNumbers;
	}

	public void setRowNumbers(List<Integer> rowNumbers) {
		this.rowNumbers = rowNumbers;
	}
}
