package com.qorql.texttoexcel.response;

import java.util.List;

import com.qorql.texttoexcel.Patient;

public class Response {
	private List<Patient> patients;
	private List<RowNumberResponse> rowNumbers;

	public List<Patient> getPatients() {
		return patients;
	}

	public void setPatients(List<Patient> patients) {
		this.patients = patients;
	}

	public List<RowNumberResponse> getRowNumbers() {
		return rowNumbers;
	}

	public void setRowNumbers(List<RowNumberResponse> rowNumbers) {
		this.rowNumbers = rowNumbers;
	}

}
