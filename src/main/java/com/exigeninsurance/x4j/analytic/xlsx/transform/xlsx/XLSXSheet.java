/*
 * Copyright 2008-2013 Exigen Insurance Solutions, Inc. All Rights Reserved.
 *
*/


package com.exigeninsurance.x4j.analytic.xlsx.transform.xlsx;

import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xssf.model.CommentsTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;


public class XLSXSheet extends XSSFSheet {

	public XLSXSheet() {}

	protected XLSXSheet(PackagePart part) {
        super(part);
    }

	@Override
	public CommentsTable getCommentsTable(boolean create) {	
		return super.getCommentsTable(create);
	}

}
