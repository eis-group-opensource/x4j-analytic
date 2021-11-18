/*
 * Copyright 2008-2020 Exigen Insurance Solutions, Inc. All Rights Reserved.
 *
*/
package com.exigeninsurance.x4j.analytic.xlsx.transform.xlsx;

import java.io.IOException;
import java.util.List;
import java.util.zip.ZipEntry;

import org.apache.commons.compress.archivers.zip.ZipArchiveEntry;
import org.apache.commons.compress.archivers.zip.ZipArchiveOutputStream;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackageProperties;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.openxml4j.opc.ZipPackage;
import org.apache.poi.openxml4j.opc.internal.ZipHelper;
import org.apache.poi.openxml4j.opc.internal.marshallers.ZipPackagePropertiesMarshaller;
import org.apache.poi.openxml4j.opc.internal.marshallers.ZipPartMarshaller;
import org.apache.poi.xssf.model.MapInfo;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.ConnectionsDocument;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.MapInfoDocument;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.WorkbookDocument;


final class XLSXMarshaler {

	private final ZipPackage pack;
	private final XSSFWorkbook workBook;
	private final ZipPartMarshaller zipPartMarshaller = new ZipPartMarshaller();

	private static final XmlOptions DEFAULT_XML_OPTIONS;

	static {
		DEFAULT_XML_OPTIONS = new XmlOptions();
		DEFAULT_XML_OPTIONS.setSaveOuter();
		DEFAULT_XML_OPTIONS.setUseDefaultNamespace();
		DEFAULT_XML_OPTIONS.setSaveAggressiveNamespaces();
	}

	public XLSXMarshaler(ZipPackage pack,
			XSSFWorkbook workBook){
		this.pack = pack;
		this.workBook  = workBook;
		

	}


	public void marshalParts(List<String> savedParts,
							 ZipArchiveOutputStream out) throws IOException,
            OpenXML4JException {

		marshalRelationship(out);

		if(workBook.getCalculationChain() != null){
			saveCalculationChain(savedParts, workBook);
		}

		saveStyles(savedParts, workBook, out);


		
		MapInfo info = null;
		XLSXConnections connections = null;

		for(POIXMLDocumentPart p : workBook.getRelations()){
			if(p instanceof MapInfo) {
				info = (MapInfo) p;
				
			}   else if (p instanceof XLSXConnections){					
				connections = ((XLSXConnections) p);
			}               

		}

		if(info != null){
			MapInfoDocument doc = MapInfoDocument.Factory.newInstance();
			doc.setMapInfo(info.getCTMapInfo());
			String entryName = getEntryName(info.getPackagePart());	
			savedParts.add(entryName);
			out.putArchiveEntry(new ZipArchiveEntry(entryName));
			doc.save(out, new XmlOptions(DEFAULT_XML_OPTIONS));
			out.closeArchiveEntry();
		}
		if(connections != null){
			ConnectionsDocument doc = ConnectionsDocument.Factory.newInstance();
			doc.setConnections(connections.getCTConnections());
			String entryName = getEntryName(connections.getPackagePart());	
			savedParts.add(entryName);
			out.putArchiveEntry(new ZipArchiveEntry(entryName));
			doc.save(out, new XmlOptions(DEFAULT_XML_OPTIONS));
			out.closeArchiveEntry();
		}


		marshalUnsavedParts(savedParts, out);



	}

    private void marshalRelationship(ZipArchiveOutputStream out) {
		ZipPartMarshaller.marshallRelationshipPart(
				pack.getRelationships(),
				PackagingURIHelper.PACKAGE_RELATIONSHIPS_ROOT_PART_NAME,
				out
		);
	}

	private void marshalUnsavedParts(List<String> savedParts,
									 ZipArchiveOutputStream out) throws
            OpenXML4JException, IOException {



		for(PackagePart part : pack.getParts() ){


			String entryName = getEntryName(part);

			if (part.isRelationshipPart()){
				continue;
			}

			if(entryName.equals(getEntryName(workBook.getSharedStringSource().getPackagePart()))){					
				continue;
			}

            if (entryName.equals(getEntryName(workBook.getPackagePart()))) {
                continue;
            }

			if(savedParts.contains(entryName)){
				continue;

			}
			if(part instanceof PackageProperties){
				new ZipPackagePropertiesMarshaller().marshall(
						(PackagePart) pack.getPackageProperties(), out);

			}else {

				zipPartMarshaller.marshall(part, out);
			}

			savedParts.add(entryName);
		}

       saveWorkBook(savedParts, workBook, out);
	}



	



	private void saveStyles(List<String> savedParts, XSSFWorkbook workBook,
							ZipArchiveOutputStream out) throws IOException {

		String entryName = getEntryName(workBook.getStylesSource().getPackagePart());
		out.putArchiveEntry(new ZipArchiveEntry(entryName));
		workBook.getStylesSource().writeTo(out);
		 out.closeArchiveEntry();

		savedParts.add(entryName);
	}

    private void saveWorkBook(List<String> savedParts, XSSFWorkbook workBook, ZipArchiveOutputStream out) throws IOException, OpenXML4JException {
        removePrintAreas(workBook);

        PackagePart part = workBook.getPackagePart();
        String entryName = getEntryName(part);
        WorkbookDocument doc = WorkbookDocument.Factory.newInstance();
        doc.setWorkbook(workBook.getCTWorkbook());
        out.putArchiveEntry(new ZipArchiveEntry(entryName));
        doc.save(out, new XmlOptions(DEFAULT_XML_OPTIONS));
        out.closeArchiveEntry();
        savedParts.add(entryName);

        // Saving relationship part
        if (part.hasRelationships()) {
            PackagePartName relationshipPartName = PackagingURIHelper
                    .getRelationshipPartName(part.getPartName());

            ZipPartMarshaller.marshallRelationshipPart(part.getRelationships(),
                    relationshipPartName, out);

            savedParts.add(getEntryName(part));

        }
    }

    private void removePrintAreas(XSSFWorkbook workBook) {
        for (int s = 0; s < workBook.getNumberOfSheets(); s++) {
            workBook.removePrintArea(s);
        }
    }

    private void saveCalculationChain(
            List<String> savedParts,
            XSSFWorkbook workBook) {

		String entryName = getEntryName(workBook.getCalculationChain().getPackagePart());	
		savedParts.add(entryName);

	}

	private String getEntryName(PackagePart packagePart) {
		return ZipHelper.getZipItemNameFromOPCName(packagePart.getPartName().getName());

	}



}
