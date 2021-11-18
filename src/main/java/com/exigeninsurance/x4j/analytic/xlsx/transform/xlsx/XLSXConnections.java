/*
 * Copyright 2008-2013 Exigen Insurance Solutions, Inc. All Rights Reserved.
 *
*/


package com.exigeninsurance.x4j.analytic.xlsx.transform.xlsx;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTConnections;


public class XLSXConnections extends POIXMLDocumentPart {
	
	private CTConnections connections;

	public XLSXConnections()  {
			
	}
	
	public XLSXConnections(PackagePart part){
		super(part);
	}

	public XLSXConnections(POIXMLDocumentPart parent, PackagePart part){
		super(part);
	}
	
	public CTConnections getCTConnections(){
		
		if(connections == null){
			connections = CTConnections.Factory.newInstance();
		}
		return connections;
		
	}

}
