package com.qorql.texttoexcel.exception;

public class TextReaderException extends Exception {

	private String message = null;

	public TextReaderException() {
		super();
	}

	public TextReaderException(String message) {
		super(message);
		this.message = message;
	}

	public TextReaderException(Throwable cause) {
		super(cause);
	}

	@Override
	public String toString() {
		return message;
	}

	@Override
	public String getMessage() {
		return message;
	}

}
