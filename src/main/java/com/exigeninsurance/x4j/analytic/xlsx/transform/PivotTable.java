/*
 * Copyright 2008-2013 Exigen Insurance Solutions, Inc. All Rights Reserved.
 *
*/


package com.exigeninsurance.x4j.analytic.xlsx.transform;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableDefinition;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.PivotTableDefinitionDocument;


public class PivotTable extends POIXMLDocumentPart {

	private CTPivotTableDefinition ctPivotTable;
	private static final XmlOptions DEFAULT_XML_OPTIONS;

	static {
		DEFAULT_XML_OPTIONS = new XmlOptions();
		DEFAULT_XML_OPTIONS.setSaveOuter();
		DEFAULT_XML_OPTIONS.setUseDefaultNamespace();
		DEFAULT_XML_OPTIONS.setSaveAggressiveNamespaces();
	}

	public PivotTable() {
		setCtPivotTable(CTPivotTableDefinition.Factory.newInstance());
	}

	public PivotTable(PackagePart part)
			throws IOException {
		super(part);
		readFrom(part.getInputStream());
	}

	public PivotTable(POIXMLDocumentPart parent, PackagePart part)
			throws IOException {
		super(parent, part);
		readFrom(part.getInputStream());
	}

	public void setCtPivotTable(CTPivotTableDefinition ctPivotTable) {
		this.ctPivotTable = ctPivotTable;
		
	}

	public CTPivotTableDefinition getCtPivotTable() {
		return ctPivotTable;
	}

	public void readFrom(InputStream is) throws IOException {
		try {
			PivotTableDefinitionDocument doc = PivotTableDefinitionDocument.Factory.parse(is);
			setCtPivotTable(doc.getPivotTableDefinition());
		} catch (XmlException e) {
			throw new IOException(e);
		}
	}

	public XSSFSheet getXSSFSheet(){
		return (XSSFSheet) getParent();
	}

	public void writeTo(OutputStream out) throws IOException {
		PivotTableDefinitionDocument doc = PivotTableDefinitionDocument.Factory.newInstance();
		doc.setPivotTableDefinition(ctPivotTable);
		doc.save(out, new XmlOptions(DEFAULT_XML_OPTIONS));
	}

	@Override
	protected void commit() throws IOException {
		PackagePart part = getPackagePart();
		OutputStream out = part.getOutputStream();
		try{
			writeTo(out);
		}finally{
			out.close();
		}
	}

	public CellReference getStartCellReference() {
		
		return new CellReference(ctPivotTable.getLocation().getRef().split(":")[0]);
	}
}
