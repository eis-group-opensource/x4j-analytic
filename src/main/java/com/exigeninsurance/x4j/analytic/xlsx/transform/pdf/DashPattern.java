/*
 * Copyright 2008-2013 Exigen Insurance Solutions, Inc. All Rights Reserved.
 *
*/


package com.exigeninsurance.x4j.analytic.xlsx.transform.pdf;

import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.BorderStyle;


public class DashPattern {

    public static final DashPattern SOLID_LINE = new DashPattern(new float[0], 0);
    public static final DashPattern DASH_DOT = new DashPattern(new float[]{3, 3, 1, 3}, 0);
    public static final DashPattern DASH_DOT_DOT = new DashPattern(new float[]{3, 2, 1, 2, 1, 2}, 0);
    public static final DashPattern DASHED = new DashPattern(new float[]{2, 1}, 0);
    public static final DashPattern DOUBLE = new DashPattern(new float[0], 0);
    public static final DashPattern DOTTED = new DashPattern(new float[]{Border.NARROW}, 0);
    public static final DashPattern SLANTED_DASH_DOT = new DashPattern(new float[0], 0);
    public static final DashPattern HAIR = new DashPattern(new float[]{1}, 0);

    private static final Map<BorderStyle, DashPattern> BORDER_PATTERNS;

    static {
        BORDER_PATTERNS = new HashMap<BorderStyle, DashPattern>();
        BORDER_PATTERNS.put(BorderStyle.DOTTED, DOTTED);
        BORDER_PATTERNS.put(BorderStyle.THIN, SOLID_LINE);
        BORDER_PATTERNS.put(BorderStyle.DASH_DOT, DASH_DOT);
        BORDER_PATTERNS.put(BorderStyle.DASH_DOT_DOT, DASH_DOT_DOT);
        BORDER_PATTERNS.put(BorderStyle.DASHED, DASHED);
        BORDER_PATTERNS.put(BorderStyle.MEDIUM, SOLID_LINE);
        BORDER_PATTERNS.put(BorderStyle.MEDIUM_DASH_DOT, DASH_DOT);
        BORDER_PATTERNS.put(BorderStyle.MEDIUM_DASH_DOT_DOT, DASH_DOT_DOT);
        BORDER_PATTERNS.put(BorderStyle.MEDIUM_DASHED, DASHED);
        BORDER_PATTERNS.put(BorderStyle.THICK, SOLID_LINE);
        BORDER_PATTERNS.put(BorderStyle.HAIR, HAIR);

        BORDER_PATTERNS.put(BorderStyle.SLANTED_DASH_DOT, SLANTED_DASH_DOT);
        BORDER_PATTERNS.put(BorderStyle.DOUBLE, DOUBLE);
    }

    private float[]dashArray;
    private int dashPhase;

    public DashPattern(float[] dashArray, int dashPhase) {
        this.dashArray = dashArray;
        this.dashPhase = dashPhase;
    }

    public float[] getDashArray() {
        return dashArray;
    }

    public int getDashPhase() {
        return dashPhase;
    }

    public static DashPattern getDashPattern(BorderStyle cellStyle) {
        if (!BORDER_PATTERNS.containsKey(cellStyle)) {
            return SOLID_LINE;
        }
        return BORDER_PATTERNS.get(cellStyle);
    }

    @Override
    public String toString() {
        return Arrays.toString(dashArray) + " " + dashPhase;
    }
}
