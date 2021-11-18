/*
 * Copyright 2008-2013 Exigen Insurance Solutions, Inc. All Rights Reserved.
 *
*/


package com.exigeninsurance.x4j.analytic.xlsx.core.expression;

import com.exigeninsurance.x4j.analytic.xlsx.transform.xlsx.XLXContext;
import org.apache.poi.xssf.usermodel.XSSFCell;

public final class NoOpExpression implements XLSXExpression {
	private final XSSFCell cell;

	public NoOpExpression(XSSFCell cell) {
		this.cell = cell;
	}

	public Object evaluate(XLXContext context) {
		switch (cell.getCellType()) {
			case BLANK:
				return "";
			case NUMERIC:
				return cell.getNumericCellValue();
			case STRING:
				return cell.getStringCellValue();
			case FORMULA:
				return cell.getCTCell().getV();
			case BOOLEAN:
				return cell.getBooleanCellValue();
			case ERROR:
				return cell.getErrorCellValue();
			default:
				return null;
		}
	}

	
}