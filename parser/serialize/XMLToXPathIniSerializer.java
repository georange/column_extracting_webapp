package com.sheret.spreadsheet.parser.serialize;

import java.io.PrintStream;
import java.util.HashMap;

import nu.xom.Attribute;
import nu.xom.Document;
import nu.xom.Element;
import nu.xom.Node;
import nu.xom.Text;
import nu.xom.converters.DOMConverter;

/* This is a utility class that converts XML to XPath INI-style format
 * which is much easier to parse by JTWin. 
 * 
 * The rules are simple:
 * <xpath>=<value>
 * 
 * Where:
 * <xpath> is a statement of the form:
 *   /x/y/z[1]
 *   Where we use the ordinal qualifier ([n]) where there are more 
 *   than one sibling with the same path. 
 *   Where only one sibling exists, we use no ordinal qualifier.
 *   
 *   Attributes can be serialized as '/<xpath>@attribute=<value>' by 
 *   changing the default configuration.
 *   
 * <value> is the value (text) found in the node. It can be empty.
 * 
 * In both cases any '=' charaters are escaped ("\=") as are any literal 
 * '\' characters ('\\'). 
 * 
 * So the value "hello=there" would be represented as: "hello\=there"
 * and the value "hello\there" would be represented as: "hello\\there"
 * 
 * We will remove any \n and replace them with spaces.
 * 
 * Any leading or trailing space surrounding the equals sign 
 * should also be ignored.
 * 
 */
public class XMLToXPathIniSerializer {

	public static final String FEATURE_WRITE_ATTRIBUTES = "writeAttributes";
	public static final String FEATURE_WRITE_ROOT_ELEMENT = "writeRootElement";
	
	boolean writeAttributes = false;
	boolean writeRootElement = false;
	
	public void setFeature(String name, boolean value) {
		if (FEATURE_WRITE_ATTRIBUTES.equals(name)) {
			writeAttributes = value;
		} else if (FEATURE_WRITE_ROOT_ELEMENT.equals(name)) {
			writeRootElement = value;
		}
	}
	
	/**
	 * Serializes a nu.xom.Document to the output stream in the XPath INI
	 * format.
	 * 
	 * @param document
	 *            the document to serialize.
	 * @param theRootPath
	 *            the element root to serialize from.
	 * @param theOutput
	 *            the output print stream
	 * @throws Exception
	 */
	public void writeToStream(Document document, String theRootPath, PrintStream theOutput) throws Exception {
		writeToStream(document.getRootElement(), theRootPath, theOutput);
	}

	/**
	 * Serializes an org.w3c.dom.Document to the output stream in the XPath INI
	 * format.
	 * 
	 * @param w3cDoc
	 *            the document to serialize.
	 * @param theRootPath
	 *            the element root to serialize from.
	 * @param theOutput
	 *            the output print stream
	 * @throws Exception
	 */
	public void writeToStream(org.w3c.dom.Document w3cDoc, String theRootPath, PrintStream theOutput) throws Exception {
		Document xomDoc = DOMConverter.convert(w3cDoc);
		writeToStream(xomDoc, theRootPath, theOutput);
	}
	
	/**
	 * Serializes a nu.xom.Node to the output stream in the XPath INI format.
	 * 
	 * @param theNode
	 *            the Xom node to serialize
	 * @param theRootPath
	 *            the element root to serialize from.
	 * @param theOutput
	 *            the output print stream
	 * @throws Exception
	 */
	public void writeToStream(Node theNode, String theRootPath, PrintStream theOutput) throws Exception {
		writeToStream(theNode, theRootPath, theOutput, new HashMap<String, Integer>(), writeRootElement);
	}

	
	/**
	 * Serializes a nu.xom.Node to the output stream in the XPath INI format.
	 * 
	 * @param theNode
	 *            the Xom node to serialize
	 * @param theRootPath
	 *            the element root to serialize from.
	 * @param theOutput
	 *            the output print stream
	 * @param whichNodeAmI
	 * @param writeRoot
	 * @throws Exception
	 */
	private void writeToStream(Node theNode, String theRootPath, PrintStream theOutput, HashMap<String, Integer> whichNodeAmI, boolean writeRoot) throws Exception {
		// Process each child of the node. Right now we'll assume that each node
		// has only two types of kids that we care about.
		// Text and other nodes...

		if (whichNodeAmI == null) {
			whichNodeAmI = new HashMap<String, Integer>();
		}

		// If this an element node, we process it.
		if (theNode instanceof Element) {
			Element theElement = (Element) theNode;

			// Which sibling of these nodes am I?
			int count = 0;
			boolean othersLikeMe = false;

			// Make the new path.
			String theNewPath = theRootPath;
			
			if (writeRoot) {
				theNewPath += "/" + theElement.getLocalName();
			}
			
			if (whichNodeAmI.containsKey(theNewPath)) {
				Integer theCount = whichNodeAmI.get(theNewPath);
				count = theCount.intValue();
				othersLikeMe = true;
			} else {
				othersLikeMe = siblingsLikeMe(theElement);
			}
			
			
			count++;
			whichNodeAmI.put(theNewPath, new Integer(count));

			if (othersLikeMe && writeRoot)
				theNewPath += "[" + Integer.toString(count) + "]";

			// serialize attributes

			if (writeAttributes && writeRoot) {
				for (int attrIndex = 0; attrIndex < theElement.getAttributeCount(); attrIndex++) {

					Attribute attr = theElement.getAttribute(attrIndex);
					theOutput
							.println(theNewPath.trim() + "@" + attr.getLocalName() + "=" + escapeData(attr.getValue()));
				}
			}

			// If we are an empty node, print it out.
			if (allEmptyChildren(theElement) && writeRoot) {
				theOutput.println(theNewPath.trim() + "=");
			}

			for (int i = 0; i < theNode.getChildCount(); i++) {
				Node theChild = theNode.getChild(i);
				writeToStream(theChild, theNewPath, theOutput, whichNodeAmI, true);
			}
		}

		// If this is a text node, we print it out.
		else if (theNode instanceof Text && writeRoot) {
			String value = theNode.getValue();
			if (theRootPath.trim().length() > 0 && value.trim().length() > 0)
				theOutput.println(theRootPath.trim() + "=" + escapeData(value));
		}
	}

	private boolean allEmptyChildren(Element theElement) throws Exception {
		// Do we have any siblings like us?
		for (int j = 0; j < theElement.getChildCount(); j++) {
			Node theChild = theElement.getChild(j);

			// If we find a text node with an actual value, we are not empty.
			if (theChild instanceof Text) {
				String theValue = theChild.getValue();
				if (theValue != null && theValue.trim().length() > 0)
					return (false);
			}

			// If we find an element child we are not empty.
			if (theChild instanceof Element)
				return (false);
		}
		return (true);
	}

	private boolean siblingsLikeMe(Element theElement) throws Exception {
		Node theParent = theElement.getParent();

		// Do we have any siblings like us?
		for (int j = 0; j < theParent.getChildCount(); j++) {
			if ((theParent.getChild(j) instanceof Element) && ((Element) theParent.getChild(j)) != theElement
					&& ((Element) theParent.getChild(j)).getLocalName().equals(theElement.getLocalName()))
				return (true);

		}
		return (false);
	}

	private String escapeData(String theString) throws Exception {
		String theResult = new String(theString.trim());
		theResult = theResult.replace("\n", " ");
		theResult = theResult.replace("\\", "\\\\");
		theResult = theResult.replace("=", "\\=");
		return (theResult);
	}
}
