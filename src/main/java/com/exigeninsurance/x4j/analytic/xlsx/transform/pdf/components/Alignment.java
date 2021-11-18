/*
 * Copyright 2008-2013 Exigen Insurance Solutions, Inc. All Rights Reserved.
 *
*/


package com.exigeninsurance.x4j.analytic.xlsx.transform.pdf.components;

import org.apache.poi.ss.usermodel.HorizontalAlignment;

public enum Alignment {
    LEFT(HorizontalAlignment.LEFT),
	CENTER(HorizontalAlignment.CENTER),
	RIGHT(HorizontalAlignment.RIGHT);

	private final HorizontalAlignment excel;

	Alignment(HorizontalAlignment excel) {
		this.excel = excel;
	}

	public HorizontalAlignment getExcel() {
		return excel;
	}
}
