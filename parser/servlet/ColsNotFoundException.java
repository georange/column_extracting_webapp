package com.sheret.spreadsheet.parser.servlet;

public class ColsNotFoundException extends RuntimeException {
	private static final long serialVersionUID = 1L;

	public ColsNotFoundException(String errorMessage) {
	      super(errorMessage);
	  }
}
