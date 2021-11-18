package com.exigeninsurance.x4j.analytic.xlsx.utils;


import com.exigeninsurance.x4j.analytic.xlsx.transform.PivotTable;
import com.exigeninsurance.x4j.analytic.xlsx.transform.PivotTableCache;
import com.exigeninsurance.x4j.analytic.xlsx.transform.xlsx.XLSXConnections;
import com.exigeninsurance.x4j.analytic.xlsx.transform.xlsx.XLSXSheet;
import com.exigeninsurance.x4j.analytic.xlsx.transform.xlsx.XLSXTheme;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ooxml.POIXMLRelation;
import org.apache.poi.poifs.crypt.dsig.DSigRelation;
import org.apache.poi.xdgf.usermodel.XDGFRelation;
import org.apache.poi.xslf.usermodel.XSLFRelation;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xwpf.usermodel.XWPFRelation;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.HashMap;
import java.util.Map;

public class XLSXDescriptorUtils {

    private static final Logger LOG = LoggerFactory.getLogger(XLSXDescriptorUtils.class);

    public static final POIXMLRelation THEME = new POIXMLRelation("application/vnd.openxmlformats-officedocument.theme+xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "/xl/theme1.xml"){};
    public static final POIXMLRelation PIVOT = new POIXMLRelation("application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable", "/xl/pivotTables/pivotTable#.xml"){};
    public static final POIXMLRelation PIVOT_CACHE = new POIXMLRelation("application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition", "/xl/pivotCache/pivotCacheDefinition#.xml"){};
    public static final POIXMLRelation CONNECTIONS = new POIXMLRelation("application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml","http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections", "/xl/connections.xml"){};
    public static final POIXMLRelation WORKSHEET = new POIXMLRelation("application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "/xl/worksheets/sheet#.xml"){};
    public static final POIXMLRelation TABLE = new POIXMLRelation("application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table", "/xl/tables/table#.xml") {};

    private static Map<String, POIXMLDocumentPart> descriptors = new HashMap<>();

    private XLSXDescriptorUtils() {}

    /*XDGFRelation (org.apache.poi.xdgf.usermodel)
      XWPFRelation (org.apache.poi.xwpf.usermodel)
      XSLFRelation (org.apache.poi.xslf.usermodel)
      DSigRelation (org.apache.poi.poifs.crypt.dsig)
      XSSFRelation (org.apache.poi.xssf.usermodel)
    */
    public static POIXMLDocumentPart getDocumentPartInstance(POIXMLRelation descriptor) {
        init();
        POIXMLRelation poixmlRelation = null;
        POIXMLDocumentPart poixmlDocumentPart = descriptors.get(descriptor.getRelation());
        if (poixmlDocumentPart == null
                && ((poixmlRelation = XSSFRelation.getInstance(descriptor.getRelation())) != null
                    || (poixmlRelation = XDGFRelation.getInstance(descriptor.getRelation())) != null
                    || (poixmlRelation = XWPFRelation.getInstance(descriptor.getRelation())) != null
                    || (poixmlRelation = XSLFRelation.getInstance(descriptor.getRelation())) != null
                    || (poixmlRelation = DSigRelation.getInstance(descriptor.getRelation())) != null)
                && poixmlRelation.getNoArgConstructor() != null) {
            poixmlDocumentPart = poixmlRelation.getNoArgConstructor().init();
        }
        LOG.debug("poixmlDocumentPart set value {}", poixmlDocumentPart);
        return poixmlDocumentPart;
    }

    private static void init() {
        LOG.debug("Init descriptors map");
        descriptors.put(THEME.getRelation(), new XLSXTheme());
        descriptors.put(PIVOT.getRelation(), new PivotTable());
        descriptors.put(PIVOT_CACHE.getRelation(), new PivotTableCache());
        descriptors.put(CONNECTIONS.getRelation(), new XLSXConnections());
        descriptors.put(WORKSHEET.getRelation(), new XLSXSheet());
        descriptors.put(TABLE.getRelation(), new XSSFTable());
    }

}
