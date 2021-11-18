/*
 * Copyright 2008-2013 Exigen Insurance Solutions, Inc. All Rights Reserved.
 *
*/


package com.exigeninsurance.x4j.analytic.xlsx.transform.pdf.components.cell;


import static com.exigeninsurance.x4j.analytic.xlsx.transform.pdf.PdfRenderer.ROW_MARGIN;
import static org.apache.poi.ss.usermodel.BorderStyle.NONE;

import java.awt.Color;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;

import com.exigeninsurance.x4j.analytic.xlsx.transform.pdf.Border;
import com.exigeninsurance.x4j.analytic.xlsx.transform.pdf.DashPattern;
import com.exigeninsurance.x4j.analytic.xlsx.transform.pdf.PdfCellNode;
import com.exigeninsurance.x4j.analytic.xlsx.transform.pdf.PdfContext;
import com.exigeninsurance.x4j.analytic.xlsx.transform.pdf.PdfRenderer;
import com.exigeninsurance.x4j.analytic.xlsx.transform.pdf.RenderingContext;
import com.exigeninsurance.x4j.analytic.xlsx.transform.pdf.RenderingParameter;
import com.exigeninsurance.x4j.analytic.xlsx.transform.pdf.components.Alignment;
import com.exigeninsurance.x4j.analytic.xlsx.transform.pdf.components.Renderer;
import com.exigeninsurance.x4j.analytic.xlsx.transform.pdf.geometry.Line;
import com.exigeninsurance.x4j.analytic.xlsx.transform.pdf.geometry.Rectangle;
import com.exigeninsurance.x4j.analytic.xlsx.utils.ColorHelper;


public abstract class AbstractCellRenderer implements Renderer {
    protected final PdfCellNode node;
    protected Rectangle drawingArea;
    protected Rectangle textArea;
    protected Rectangle fillArea;
    private VerticalOffsetCalculator calculator = new VerticalOffsetCalculator();

    public AbstractCellRenderer(PdfCellNode node) {
        this.node = node;
    }

    public final void render(RenderingContext context) throws IOException {
		Object value = context.getParams().get(RenderingParameter.CELL_VALUE);
        node.determineAndSetAlignment(context.getPdfContext(), value);
        drawingArea = calculateDrawingArea(context,value);
        textArea = drawingArea.expand(-context.getPdfContext().getMargins(), 0);
        fillArea = calculateFillArea(context, drawingArea);
        fill(context);
        
        drawText(context, value);
        drawBorders(context);
        
    }
    
    public float findVerticalOffset(RenderingContext context, List<String> items) {
        return calculator.calculate(node.getVerticalAlignment(),
                items.size(),
                getRowHeight(context),
                node.getMaxFontHeight(),
                PdfRenderer.ROW_MARGIN);
	}

    protected abstract void drawText(RenderingContext context, Object value) throws IOException;

    protected abstract boolean applyFillAndBorders(RenderingContext context);

    protected Rectangle calculateDrawingArea(RenderingContext context, Object value) {
		PdfContext pdfContext = context.getPdfContext();
		float x = pdfContext.getX();
        float y = pdfContext.getY();
        float width = node.estimateWidth(context);
        float height = getRowHeight(context);
        return new Rectangle(x, y, width, height);
    }

	protected Float getRowHeight(RenderingContext context) {
		return (Float) context.getParams().get(RenderingParameter.ROW_HEIGHT);
	}

	protected Rectangle calculateFillArea(RenderingContext context, Rectangle drawingArea) {
        return applyFillAndBorders(context) ? drawingArea.expand(0, ROW_MARGIN) : null;
    }

    protected void fill(RenderingContext context) throws IOException {
        if (applyFillAndBorders(context)) {
            context.getPdfContext().fill(fillArea, node.getFillColor());
        }
    }

    protected void drawBorders(RenderingContext context) {
        if (applyFillAndBorders(context)) {
            drawBorders(context.getPdfContext(), node.getCellStyle());
        }
    }

    protected void drawBorders(PdfContext context, XSSFCellStyle cellStyle) {
        drawTop(context, cellStyle);
        drawBottom(context, cellStyle);
        drawLeft(context, cellStyle);
        drawRight(context, cellStyle);
    }

    private void drawRight(PdfContext context, XSSFCellStyle cellStyle) {
        if (cellStyle.getBorderRightEnum() != NONE) {
            context.drawBorder(createBorder(fillArea, XSSFCellBorder.BorderSide.RIGHT, cellStyle));
        }
    }

    private void drawLeft(PdfContext context, XSSFCellStyle cellStyle) {
        if (cellStyle.getBorderLeftEnum() != NONE) {
            context.drawBorder(createBorder(fillArea, XSSFCellBorder.BorderSide.LEFT, cellStyle));
        }
    }

    private void drawBottom(PdfContext context, XSSFCellStyle cellStyle) {
        if (cellStyle.getBorderBottomEnum() != NONE) {
            context.drawBorder(createBorder(fillArea, XSSFCellBorder.BorderSide.BOTTOM, cellStyle));
        }
    }

    private void drawTop(PdfContext context, XSSFCellStyle cellStyle) {
        if (cellStyle.getBorderTopEnum() != NONE) {
            context.drawBorder(createBorder(fillArea, XSSFCellBorder.BorderSide.TOP, cellStyle));
        }
    }

    protected Border createBorder(Rectangle fillArea, XSSFCellBorder.BorderSide borderSide, XSSFCellStyle cellStyle) {
        DashPattern pattern = getBorderPattern(cellStyle, borderSide);
        float width = getBorderWidth(getBorderStyle(cellStyle, borderSide));
        Line line = getBorderLine(fillArea, borderSide);
        ColorHelper helper = node.getColorHelper();
        XSSFColor borderColor = cellStyle.getBorderColor(borderSide);
        Color color = helper == null || borderColor.isAuto() ? Color.BLACK : helper.getAwtColor(borderColor);
        return new Border(color, line, pattern, width);
    }

    private DashPattern getBorderPattern(XSSFCellStyle cellStyle, XSSFCellBorder.BorderSide borderSide) {
        return DashPattern.getDashPattern(getBorderStyle(cellStyle, borderSide));
    }

    private float getBorderWidth(BorderStyle borderStyle) {
        switch (borderStyle) {
            case HAIR:
            case THIN:
            case DOTTED:
            case DASH_DOT:
            case DASH_DOT_DOT:
            case DASHED:
            case DOUBLE:
                return Border.NARROW;
            case MEDIUM:
            case MEDIUM_DASHED:
            case MEDIUM_DASH_DOT:
            case MEDIUM_DASH_DOT_DOT:
            case SLANTED_DASH_DOT:
                return Border.MEDIUM;
            case THICK:
                return Border.WIDE;
            default :
                return Border.MEDIUM;
        }
    }

    private Line getBorderLine(Rectangle fillArea, XSSFCellBorder.BorderSide borderSide) {
        Line line = null;
        switch (borderSide) {
            case BOTTOM:
                line = fillArea.getBottom();
                break;
            case LEFT:
                line = fillArea.getLeft();
                break;
            case RIGHT:
                line = fillArea.getRight();
                break;
            case TOP:
                line = fillArea.getTop();
                break;
        }
        return line;
    }

    private BorderStyle getBorderStyle(XSSFCellStyle cellStyle, XSSFCellBorder.BorderSide borderSide) {
        switch (borderSide) {
            case BOTTOM:
                return cellStyle.getBorderBottomEnum();
            case LEFT:
                return cellStyle.getBorderLeftEnum();
            case RIGHT:
                return cellStyle.getBorderRightEnum();
            case TOP:
                return cellStyle.getBorderTopEnum();
        }
        throw new IllegalStateException();
    }

    protected float findHorizontalOffset(String value) {
        Alignment alignment = node.getHorizontalAlignment();
        float textLength = node.findTextLength(value);
        switch (alignment) {
            case CENTER:
                return -(textLength / 2);

            case RIGHT:
                return -textLength;

            default:
                return 0f;
        }
    }

    protected float findStartingDrawingPoint(float cellWidth) {
        Alignment alignment = node.getHorizontalAlignment();
        switch (alignment) {
            case CENTER:
                return cellWidth * 0.5f;

            case RIGHT:
                return cellWidth;
            default:
                return 0f;
        }
    }

    protected float findVerticalOffset(float containerHeight) {
        VerticalAlignment alignment = node.getVerticalAlignment();
        float fontHeight = node.getMaxFontHeight();

        switch (alignment) {
            case CENTER:
                return (containerHeight - fontHeight) / 2;
            case TOP:
                return containerHeight - fontHeight;

            default:
                return 0f;

        }
    }

	public void setTextOptions(PdfContext context) throws IOException {
		context.setTextOptions(node.findFontSize(), node.getFont(), node.getUnderline());
	}
}
