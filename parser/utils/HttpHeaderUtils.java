package com.sheret.spreadsheet.parser.utils;

import java.io.UnsupportedEncodingException;

import javax.servlet.http.Part;

import org.apache.http.Header;

public class HttpHeaderUtils
{
	public static final String HTTP_HEADER_CONTENT_DISP = "content-disposition";
	
	public static String getFilenameFromPart(Part part) throws Exception
	{
		String result = null;
		String cdHeaderValue = part.getHeader(HTTP_HEADER_CONTENT_DISP);

		if (cdHeaderValue != null) {
			result = getFilenameFromHeaderValue(cdHeaderValue);
		}

		return result;
	}

	public static String getFilenameFromHeader(Header header) throws Exception
	{
		String result = null;

		if (header != null) {
			String headerValue = header.getValue();

			result = getFilenameFromHeaderValue(headerValue);
		}

		return result;
	}

	public static String getFilenameFromHeaderValue(String headerValue) throws Exception
	{
		String result = null;

		if ((headerValue != null) && (headerValue.trim().length() > 0))
		{
			for (String cd : headerValue.trim().split(";")) 
			{
				if (cd.trim().startsWith("filename")) 
				{
					String filename = cd.substring(cd.indexOf('=') + 1).trim().replace("\"", "");
					result = filename.substring(filename.lastIndexOf('/') + 1).substring(filename.lastIndexOf('\\') + 1); // MSIE fix.
					break;
				}
			}
		}

		return result;
	}

	public static String createContentDispositionHeaderValue(String filename) throws UnsupportedEncodingException {
		StringBuilder cdHeaderValue = new StringBuilder("attachment");

		if ((filename != null) && (filename.isEmpty() == false)) {
			cdHeaderValue.append("; filename=");
			cdHeaderValue.append(java.net.URLEncoder.encode(filename, "UTF-8"));
		}
		
		return cdHeaderValue.toString();
	}
}

