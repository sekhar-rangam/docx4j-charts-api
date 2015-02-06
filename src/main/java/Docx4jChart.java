package main.java;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.Reader;
import java.io.UnsupportedEncodingException;

import javax.xml.bind.JAXBException;

import org.docx4j.XmlUtils;
import org.docx4j.dml.chart.CTChartSpace;
import org.docx4j.openpackaging.contenttype.ContentType;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.OpcPackage;
import org.docx4j.openpackaging.packages.PresentationMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.PresentationML.MainPresentationPart;
import org.docx4j.openpackaging.parts.PresentationML.SlideLayoutPart;
import org.docx4j.openpackaging.parts.PresentationML.SlidePart;
import org.docx4j.openpackaging.parts.WordprocessingML.EmbeddedPackagePart;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.relationships.Relationships;
import org.pptx4j.jaxb.Context;
import org.pptx4j.pml.Sld;

public class Docx4jChart {

	public static void main(String[] args) throws JAXBException,
			FileNotFoundException, Docx4JException,
			UnsupportedEncodingException {

		try {
			ContentType contentType = new ContentType(
					"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
			String excelFilePath = "e://data.xlsx";
			PresentationMLPackage presentationMLPackage = null;
			int slideIndex = 11;
			String templateFile = "/main/resources/template.pptx";
			String destinationFolderStr = "c:/ppt/";
			File destinationFolder = new File(destinationFolderStr);
			if (!destinationFolder.exists()) {
				destinationFolder.mkdir();
			}
			destinationFolderStr = destinationFolderStr + File.separator
					+ "docx4JChart.pptx";
			try {
				InputStream in = Docx4jChart.class
						.getResourceAsStream(templateFile);
				presentationMLPackage = (PresentationMLPackage) OpcPackage
						.load(in);
			} catch (Exception e) {
				// TODO: handle exception
				presentationMLPackage = PresentationMLPackage.createPackage();
			}

			MainPresentationPart pp = (MainPresentationPart) presentationMLPackage
					.getParts().getParts()
					.get(new PartName("/ppt/presentation.xml"));
			SlideLayoutPart layoutPart = (SlideLayoutPart) presentationMLPackage
					.getParts().getParts()
					.get(new PartName("/ppt/slideLayouts/slideLayout2.xml"));

			SlidePart slide1 = presentationMLPackage.createSlidePart(pp,
					layoutPart, new PartName("/ppt/slides/slide" + slideIndex
							+ ".xml"));
			StringBuffer slideXMLBuffer = new StringBuffer();
			BufferedReader br = null;
			String line = "";
			String slideDataXmlFile = "/main/data/slide_data.xml";
			InputStream in = Docx4jChart.class
					.getResourceAsStream(slideDataXmlFile);
			Reader fr = new InputStreamReader(in, "utf-8");
			br = new BufferedReader(fr);
			while ((line = br.readLine()) != null) {
				slideXMLBuffer.append(line);
				slideXMLBuffer.append(" ");
			}

			Sld sld = (Sld) XmlUtils.unmarshalString(slideXMLBuffer.toString(),
					Context.jcPML, Sld.class);
			slide1.setJaxbElement(sld);

			org.docx4j.openpackaging.parts.DrawingML.Chart chartPart = new org.docx4j.openpackaging.parts.DrawingML.Chart(
					new PartName("/ppt/charts/chart" + slideIndex + ".xml"));

			StringBuffer chartXMLBuffer = new StringBuffer();
			String chartDataXmlFile = "/main/data/chart_data.xml";
			in = Docx4jChart.class.getResourceAsStream(chartDataXmlFile);
			fr = new InputStreamReader(in, "utf-8");
			br = new BufferedReader(fr);
			while ((line = br.readLine()) != null) {
				chartXMLBuffer.append(line);
				chartXMLBuffer.append(" ");
			}

			CTChartSpace chartSpace = (CTChartSpace) XmlUtils.unmarshalString(
					chartXMLBuffer.toString(), Context.jcPML,
					CTChartSpace.class);
			chartPart.setJaxbElement(chartSpace);

			slide1.addTargetPart(chartPart);
			EmbeddedPackagePart embeddedPackagePart = new EmbeddedPackagePart(
					new PartName("/ppt/embeddings/Microsoft_Excel_Worksheet"
							+ slideIndex + ".xlsx"));
			embeddedPackagePart.setContentType(contentType);
			embeddedPackagePart
					.setRelationshipType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/package");
			RelationshipsPart owningRelationshipPart = new RelationshipsPart();
			PartName partName = new PartName("/ppt/charts/_rels/chart"
					+ slideIndex + ".xml.rels");
			owningRelationshipPart.setPartName(partName);
			Relationship relationship = new Relationship();
			relationship.setId("rId1");
			relationship.setTarget("../embeddings/Microsoft_Excel_Worksheet"
					+ slideIndex + ".xlsx");
			relationship
					.setType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/package");
			owningRelationshipPart.setRelationships(new Relationships());
			owningRelationshipPart.addRelationship(relationship);
			embeddedPackagePart
					.setOwningRelationshipPart(owningRelationshipPart);
			embeddedPackagePart.setPackage(presentationMLPackage);
			String dataFile = "/main/data/data.xlsx";
			InputStream inputStream = Docx4jChart.class
					.getResourceAsStream(dataFile);
			embeddedPackagePart.setBinaryData(inputStream);
			chartPart.addTargetPart(embeddedPackagePart);
			presentationMLPackage.save(new java.io.File(destinationFolderStr));
			System.out.println("Powerpoint Office document is created in the following path: "+destinationFolderStr);
		} catch (Exception exception) {
			exception.printStackTrace();
		}
	}
}