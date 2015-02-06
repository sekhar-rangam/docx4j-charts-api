package main.java;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.Reader;

import javax.xml.bind.JAXBException;

import org.docx4j.XmlUtils;
import org.docx4j.dml.chart.CTChartSpace;
import org.docx4j.openpackaging.contenttype.ContentType;
import org.docx4j.openpackaging.exceptions.Docx4JException;
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
			FileNotFoundException, Docx4JException {
		// TODO Auto-generated method stub

		ContentType contentType = new ContentType(
				"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
		String excelFilePath = "e://data.xlsx";
		PresentationMLPackage presentationMLPackage = null;
		int slideIndex = 11;
		String pptxFilePath = "e:\\myFirstChart.pptx";
		try {
			presentationMLPackage = PresentationMLPackage.load(new File(
					pptxFilePath));
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
		try {
			InputStream in = Docx4jChart.class
					.getResourceAsStream(slideDataXmlFile);
			Reader fr = new InputStreamReader(in, "utf-8");
			br = new BufferedReader(fr);
			while ((line = br.readLine()) != null) {
				slideXMLBuffer.append(line);
				slideXMLBuffer.append(" ");
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (br != null) {
				try {
					br.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		// slideXMLBuffer.append("<p:sld xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></p:grpSpPr><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id=\"4\" name=\"Chart 3\"/><p:cNvGraphicFramePr/><p:nvPr><p:extLst><p:ext uri=\"{D42A27DB-BD31-4B8C-83A1-F6EECF244321}\"><p14:modId xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\" val=\"3565175352\"/></p:ext></p:extLst></p:nvPr></p:nvGraphicFramePr><p:xfrm><a:off x=\"228600\" y=\"2133600\"/><a:ext cx=\"8610600\" cy=\"4064000\"/></p:xfrm><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId2\"/></a:graphicData></a:graphic></p:graphicFrame><p:sp>	<p:nvSpPr>		<p:cNvPr id=\"2\" name=\"TextBox 1\" />		<p:cNvSpPr txBox=\"1\" />		<p:nvPr />	</p:nvSpPr>	<p:spPr>		<a:xfrm>			<a:off x=\"990600\" y=\"533400\" />			<a:ext cx=\"6553200\" cy=\"369332\" />		</a:xfrm>		<a:prstGeom prst=\"rect\">			<a:avLst />		</a:prstGeom>		<a:noFill />	</p:spPr>	<p:txBody>		<a:bodyPr wrap=\"square\" rtlCol=\"0\">			<a:spAutoFit />		</a:bodyPr>		<a:lstStyle />		<a:p>			<a:endParaRPr lang=\"en-US\" dirty=\"0\" />		</a:p>	</p:txBody></p:sp><p:sp>	<p:nvSpPr>		<p:cNvPr id=\"3\" name=\"Rectangle 2\" />		<p:cNvSpPr />		<p:nvPr />	</p:nvSpPr>	<p:spPr>		<a:xfrm>			<a:off x=\"0\" y=\"272534\" />			<a:ext cx=\"8382000\" cy=\"630198\" />		</a:xfrm>		<a:prstGeom prst=\"rect\">			<a:avLst />		</a:prstGeom>		<a:solidFill>			<a:schemeClr val=\"bg1\">				<a:lumMod val=\"95000\" />			</a:schemeClr>		</a:solidFill>		<a:ln>			<a:noFill />		</a:ln>	</p:spPr>	<p:style>		<a:lnRef idx=\"2\">			<a:schemeClr val=\"accent1\">				<a:shade val=\"50000\" />			</a:schemeClr>		</a:lnRef>		<a:fillRef idx=\"1\">			<a:schemeClr val=\"accent1\" />		</a:fillRef>		<a:effectRef idx=\"0\">			<a:schemeClr val=\"accent1\" />		</a:effectRef>		<a:fontRef idx=\"minor\">			<a:schemeClr val=\"lt1\" />		</a:fontRef>	</p:style>	<p:txBody>		<a:bodyPr rtlCol=\"0\" anchor=\"ctr\" />		<a:lstStyle />		<a:p>			<a:pPr algn=\"ctr\" />			<a:endParaRPr lang=\"en-US\" />		</a:p>	</p:txBody></p:sp><p:sp>	<p:nvSpPr>		<p:cNvPr id=\"5\" name=\"TextBox 4\" />		<p:cNvSpPr txBox=\"1\" />		<p:nvPr />	</p:nvSpPr>	<p:spPr>		<a:xfrm>			<a:off x=\"76200\" y=\"316468\" />			<a:ext cx=\"8305800\" cy=\"369332\" />		</a:xfrm>		<a:prstGeom prst=\"rect\">			<a:avLst />		</a:prstGeom>		<a:noFill />	</p:spPr>	<p:txBody>		<a:bodyPr wrap=\"square\" rtlCol=\"0\">			<a:spAutoFit />		</a:bodyPr>		<a:lstStyle />		<a:p>			<a:r>				<a:rPr lang=\"en-US\" dirty=\"0\" smtClean=\"0\" />				<a:t>Frequency Chart</a:t>			</a:r>			<a:endParaRPr lang=\"en-US\" dirty=\"0\" />		</a:p>	</p:txBody></p:sp></p:spTree><p:extLst><p:ext uri=\"{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}\"><p14:creationId xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\" val=\"3546847285\"/></p:ext></p:extLst></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>");
		Sld sld = (Sld) XmlUtils.unmarshalString(slideXMLBuffer.toString(),
				Context.jcPML, Sld.class);
		slide1.setJaxbElement(sld);

		org.docx4j.openpackaging.parts.DrawingML.Chart chartPart = new org.docx4j.openpackaging.parts.DrawingML.Chart(
				new PartName("/ppt/charts/chart" + slideIndex + ".xml"));

		StringBuffer chartXMLBuffer = new StringBuffer();
		String chartDataXmlFile = "/main/data/chart_data.xml";
		try {
			InputStream in = Docx4jChart.class
					.getResourceAsStream(chartDataXmlFile);
			Reader fr = new InputStreamReader(in, "utf-8");
			br = new BufferedReader(fr);
			while ((line = br.readLine()) != null) {
				chartXMLBuffer.append(line);
				chartXMLBuffer.append(" ");
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (br != null) {
				try {
					br.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}

		CTChartSpace chartSpace = (CTChartSpace) XmlUtils.unmarshalString(
				chartXMLBuffer.toString(), Context.jcPML, CTChartSpace.class);
		chartPart.setJaxbElement(chartSpace);

		slide1.addTargetPart(chartPart);
		EmbeddedPackagePart embeddedPackagePart = new EmbeddedPackagePart(
				new PartName("/ppt/embeddings/Microsoft_Excel_Worksheet"
						+ slideIndex + ".xlsx"));
		embeddedPackagePart.setContentType(contentType);
		embeddedPackagePart
				.setRelationshipType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/package");
		RelationshipsPart owningRelationshipPart = new RelationshipsPart();
		PartName partName = new PartName("/ppt/charts/_rels/chart" + slideIndex
				+ ".xml.rels");
		owningRelationshipPart.setPartName(partName);
		Relationship relationship = new Relationship();
		relationship.setId("rId1");
		relationship.setTarget("../embeddings/Microsoft_Excel_Worksheet"
				+ slideIndex + ".xlsx");
		relationship
				.setType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/package");
		owningRelationshipPart.setRelationships(new Relationships());
		owningRelationshipPart.addRelationship(relationship);
		embeddedPackagePart.setOwningRelationshipPart(owningRelationshipPart);
		embeddedPackagePart.setPackage(presentationMLPackage);
		InputStream inputStream = new FileInputStream(new File(excelFilePath));
		embeddedPackagePart.setBinaryData(inputStream);
		chartPart.addTargetPart(embeddedPackagePart);
		presentationMLPackage.save(new java.io.File(pptxFilePath));

	}

}
