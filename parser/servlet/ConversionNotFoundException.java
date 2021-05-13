package com.sheret.spreadsheet.parser.servlet;

public class ConversionNotFoundException extends RuntimeException {
	private static final long serialVersionUID = 1L;

	public ConversionNotFoundException(String errorMessage) {
	      super(errorMessage);
	  }
}
