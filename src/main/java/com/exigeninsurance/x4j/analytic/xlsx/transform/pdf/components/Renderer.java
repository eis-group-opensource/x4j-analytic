/*
 * Copyright 2008-2013 Exigen Insurance Solutions, Inc. All Rights Reserved.
 *
*/


package com.exigeninsurance.x4j.analytic.xlsx.transform.pdf.components;

import java.io.IOException;

import com.exigeninsurance.x4j.analytic.xlsx.transform.pdf.RenderingContext;


public interface Renderer {

	void render(RenderingContext context) throws IOException;
}
