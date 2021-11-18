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
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotCacheDefinition;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.PivotCacheDefinitionDocument;


public class PivotTableCache extends POIXMLDocumentPart {

	private CTPivotCacheDefinition ctPivotTable;
	private static final XmlOptions DEFAULT_XML_OPTIONS;

	static {
		DEFAULT_XML_OPTIONS = new XmlOptions();
		DEFAULT_XML_OPTIONS.setSaveOuter();
		DEFAULT_XML_OPTIONS.setUseDefaultNamespace();
		DEFAULT_XML_OPTIONS.setSaveAggressiveNamespaces();
	}

	public PivotTableCache() {
		setCtPivotTable(CTPivotCacheDefinition.Factory.newInstance());
	}

	public PivotTableCache(PackagePart part)
			throws IOException {
		super(part);
		readFrom(part.getInputStream());
	}

	public PivotTableCache(POIXMLDocumentPart parent, PackagePart part)
			throws IOException {
		super(part);
		readFrom(part.getInputStream());
	}

	public void readFrom(InputStream is) throws IOException {
		try {
			PivotCacheDefinitionDocument doc = PivotCacheDefinitionDocument.Factory.parse(is);
			setCtPivotTable( doc.getPivotCacheDefinition() );
		} catch (XmlException e) {
			throw new IOException(e);
		}
	}

	public XSSFSheet getXSSFSheet(){
		return (XSSFSheet) getParent();
	}

	public void writeTo(OutputStream out) throws IOException {
		PivotCacheDefinitionDocument doc = PivotCacheDefinitionDocument.Factory.newInstance();
		doc.setPivotCacheDefinition(ctPivotTable);
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

	public void setCtPivotTable(CTPivotCacheDefinition ctPivotTable) {
		this.ctPivotTable = ctPivotTable;
		this.ctPivotTable.setSaveData(false);
		this.ctPivotTable.setRefreshOnLoad(true);
		if(this.ctPivotTable.isSetId()){
			this.ctPivotTable.unsetId();
		}
	}

	public CTPivotCacheDefinition getCtPivotTable() {
		return ctPivotTable;
	}
}
