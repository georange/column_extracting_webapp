package com.sheret.spreadsheet.parser.servlet;

public class MissingParamException extends RuntimeException {
	private static final long serialVersionUID = 1L;

	public MissingParamException(String errorMessage) {
	      super(errorMessage);
	  }
	}