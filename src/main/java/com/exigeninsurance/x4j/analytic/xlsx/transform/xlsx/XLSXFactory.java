/*
 * Copyright 2008-2020 Exigen Insurance Solutions, Inc. All Rights Reserved.
 *
 */
package com.exigeninsurance.x4j.analytic.xlsx.transform.xlsx;

import com.exigeninsurance.x4j.analytic.xlsx.utils.XLSXDescriptorUtils;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ooxml.POIXMLException;
import org.apache.poi.ooxml.POIXMLFactory;
import org.apache.poi.ooxml.POIXMLRelation;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;

import static com.exigeninsurance.x4j.analytic.xlsx.utils.XLSXDescriptorUtils.*;

public final class XLSXFactory extends POIXMLFactory {

    private static final Logger LOG = LoggerFactory.getLogger(XLSXFactory.class);

    private static final Class<?>[] PARENT_PART = new Class[]{POIXMLDocumentPart.class, PackagePart.class};
    private static final Class<?>[] ORPHAN_PART = new Class[]{PackagePart.class};

    @Override
    public POIXMLDocumentPart createDocumentPart(POIXMLDocumentPart parent, PackagePart part) {
        LOG.debug("Create document part for parent {} and part {} started", parent, part);
        POIXMLDocumentPart poixmlDocumentPart = null;
        PackageRelationship rel = this.getPackageRelationship(parent, part);
        POIXMLRelation descriptor = this.getDescriptor(rel.getRelationshipType());
        if (descriptor != null) {
             poixmlDocumentPart = getDocumentPartInstance(descriptor);
            if (poixmlDocumentPart != null) {
                final Class<? extends POIXMLDocumentPart> cls = poixmlDocumentPart.getClass();
                try {
                    try {
                        poixmlDocumentPart = this.create(cls, PARENT_PART, new Object[]{parent, part});
                    } catch (NoSuchMethodException var7) {
                        poixmlDocumentPart = this.create(cls, ORPHAN_PART, new Object[]{part});
                    }
                } catch (Exception var8) {
                    throw new POIXMLException(var8);
                }
            } else {
                poixmlDocumentPart = new POIXMLDocumentPart(parent, part);
            }
        } else {
            poixmlDocumentPart = new POIXMLDocumentPart(parent, part);
        }
        LOG.debug("poixmlDocumentPart set value {}", poixmlDocumentPart);
        return poixmlDocumentPart;
    }

    protected POIXMLDocumentPart create(Class<? extends POIXMLDocumentPart> aClass, Class<?>[] classes, Object[] objects) throws NoSuchMethodException, InstantiationException, IllegalAccessException, InvocationTargetException {
        Class<? extends POIXMLDocumentPart> cls = aClass == StylesTable.class ? XLSXStylesTable.class : aClass;
        if (classes.length > 1) {
            Constructor<? extends POIXMLDocumentPart> constructor = cls.getDeclaredConstructor(classes[0], classes[1]);
            constructor.setAccessible(true);
            return constructor.newInstance(objects[0], objects[1]);
        } else {
            Constructor<? extends POIXMLDocumentPart> constructor = cls.getDeclaredConstructor(classes[0]);
            constructor.setAccessible(true);
            return constructor.newInstance(objects[0]);
        }
    }

    @Override
    protected POIXMLRelation getDescriptor(String s) {
        if (THEME.getRelation().equals(s)) {
            return THEME;
        } else if (PIVOT.getRelation().equals(s)) {
            return PIVOT;
        } else if (PIVOT_CACHE.getRelation().equals(s)) {
            return PIVOT_CACHE;
        } else if (CONNECTIONS.getRelation().equals(s)) {
            return CONNECTIONS;
        } else if (WORKSHEET.getRelation().equals(s)) {
            return WORKSHEET;
        }
        return XSSFRelation.getInstance(s);
    }

    @Override
    public POIXMLDocumentPart newDocumentPart(POIXMLRelation descriptor) {
        return XLSXDescriptorUtils.getDocumentPartInstance(descriptor);
    }
}
