/*
 * Copyright 2008-2013 Exigen Insurance Solutions, Inc. All Rights Reserved.
 *
*/


package com.exigeninsurance.x4j.analytic.xlsx.transform.pdf.components.cell;

import org.apache.poi.ss.usermodel.VerticalAlignment;


public class VerticalOffsetCalculator {
    public float calculate(VerticalAlignment alignment, float rowCount, float totalHeight, float rowHeight,  float margin) {
        float itemHeight = rowHeight * rowCount + margin * (rowCount - 1);
        switch (alignment) {
            case CENTER:
                return totalHeight / 2 - itemHeight / 2 ;
            case TOP:
                return totalHeight - itemHeight;

            default:
                return 0f;

        }
    }
}
