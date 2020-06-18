/*
 *  Copyright 2007-2008, Plutext Pty Ltd.
 *   
 *  This file is part of docx4j.

    docx4j is licensed under the Apache License, Version 2.0 (the "License"); 
    you may not use this file except in compliance with the License. 

    You may obtain a copy of the License at 

        http://www.apache.org/licenses/LICENSE-2.0 

    Unless required by applicable law or agreed to in writing, software 
    distributed under the License is distributed on an "AS IS" BASIS, 
    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. 
    See the License for the specific language governing permissions and 
    limitations under the License.

 */



import java.io.File;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;
import java.util.Map;

import javax.xml.bind.JAXBContext;

import org.docx4j.Docx4J;
import org.docx4j.XmlUtils;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.diff.Differencer;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart.AddPartBehaviour;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.Body;
import org.docx4j.wml.Document;

import com.topologi.diffx.Docx4jDriver;


/**
 * This sample compares the 2 input documents, and renders
 * the result using PDF viewer.
 *
 */
public class CompareDocuments {
	
	public static JAXBContext context = org.docx4j.jaxb.Context.jc; 
	
	static boolean DOCX_SAVE = true;

	static boolean PDF_SAVE = false;
	
	
	/**
	 * Split up the problem to try to solve more quickly
	 * (you might try this when you have 500 entries or more)
	 */
	static boolean DIVIDE_AND_CONQUER = false;
	
	
	/**
	 * @param args
	 */
	public static void main(String[] args) throws Exception {

		String newerfilepath = "/Users/faresyoussef/Desktop/testing differences/sample-docx.docx";
		String olderfilepath = "/Users/faresyoussef/Desktop/testing differences/sample-docxv2.docx";
		
		// 1. Load the Packages
		WordprocessingMLPackage newerPackage = WordprocessingMLPackage.load(new java.io.File(newerfilepath));
		WordprocessingMLPackage olderPackage = WordprocessingMLPackage.load(new java.io.File(olderfilepath));
		
		Body newerBody = ((Document)newerPackage.getMainDocumentPart().getJaxbElement()).getBody();
		Body olderBody = ((Document)olderPackage.getMainDocumentPart().getJaxbElement()).getBody();
		
		
		// 2. Do the differencing
		java.io.StringWriter sw = new java.io.StringWriter();
		javax.xml.transform.stream.StreamResult result = new javax.xml.transform.stream.StreamResult(
				sw);
		Calendar changeDate = null;
		
		Differencer pd = null;
		if (DIVIDE_AND_CONQUER) {

			System.out.println("Differencing with DIVIDE_AND_CONQUER..");
			
			java.io.StringWriter swInterim = new java.io.StringWriter();
			
			Docx4jDriver.diff( XmlUtils.marshaltoW3CDomDocument(newerBody).getDocumentElement(),
					XmlUtils.marshaltoW3CDomDocument(olderBody).getDocumentElement(),
					swInterim);
				// The signature which takes Reader objects appears to be broken
			
			// Now, feed it through diff to wml XSLT
			pd = new Differencer();
			pd.toWML(swInterim.toString(), result, "someone", changeDate,
					newerPackage.getMainDocumentPart().getRelationshipsPart(),
					olderPackage.getMainDocumentPart().getRelationshipsPart() 
					);
			
		} else {

			System.out.println("Differencing without dividing..");
			
			pd = new Differencer();
			pd.setRelsDiffIdentifier("blagh"); // not necessary in this case 
			pd.diff(newerBody, olderBody, result, "someone", changeDate,
					newerPackage.getMainDocumentPart().getRelationshipsPart(),
					olderPackage.getMainDocumentPart().getRelationshipsPart() 
					);
		//OKAY TILL HERE
		
		}
		
		// 3. Get the result
		String contentStr = sw.toString();
//		System.out.println("Result: \n\n " + contentStr);
		System.out.println("THE CONTENT:  \n\n" +contentStr);
		
//		String tempContent = "<?xml version=\"1.0\" encoding=\"utf-8\"?><w:body xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:dfx=\"http://www.topologi.com/2005/Diff-X\" xmlns:del=\"http://www.topologi.com/2005/Diff-X/Delete\" xmlns:ins=\"http://www.topologi.com/2005/Diff-X\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\">\n" + 
//				"    <w:p xmlns:ins=\"http://www.topologi.com/2005/Diff-X/Insert\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:pStyle w:val=\"Title\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\" xmlns:ins=\"http://www.topologi.com/2005/Diff-X\">\n" + 
//				"            <w:t xml:space=\"preserve\">Docx sample document</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">This is a document exhibiting basic docx features.</w:t>\n" + 
//				"        </w:r>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">  </w:t>\n" + 
//				"        </w:r>\n" + 
//				"        <w:del xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\" w:date=\"2009-03-11T17:57:00Z\" w:author=\"someone\" w:id=\"1\">\n" + 
//				"            <w:r>\n" + 
//				"                <w:t xml:space=\"preserve\">Compared to the Flat OPC version, it contains a few innocuous differents.</w:t>\n" + 
//				"            </w:r>\n" + 
//				"        </w:del>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:pStyle w:val=\"Heading1\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">This is style Heading 1</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00D15781\"/>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">Some text.</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:pStyle w:val=\"Heading1\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">Tables</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00D15781\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:ind w:left=\"0\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"    </w:p>\n" + 
//				"    <w:tbl>\n" + 
//				"        <w:tblPr>\n" + 
//				"            <w:tblStyle w:val=\"TableGrid\"/>\n" + 
//				"            <w:tblW w:type=\"auto\" w:w=\"0\"/>\n" + 
//				"            <w:tblLook w:firstColumn=\"1\" w:firstRow=\"1\" w:lastColumn=\"0\" w:lastRow=\"0\" w:noHBand=\"0\" w:noVBand=\"1\" w:val=\"04A0\"/>\n" + 
//				"        </w:tblPr>\n" + 
//				"        <w:tblGrid>\n" + 
//				"            <w:gridCol w:w=\"3561\"/>\n" + 
//				"            <w:gridCol w:w=\"3561\"/>\n" + 
//				"            <w:gridCol w:w=\"3561\"/>\n" + 
//				"        </w:tblGrid>\n" + 
//				"        <w:tr w:rsidR=\"00D15781\">\n" + 
//				"            <w:tc>\n" + 
//				"                <w:tcPr>\n" + 
//				"                    <w:tcW w:type=\"dxa\" w:w=\"3561\"/>\n" + 
//				"                </w:tcPr>\n" + 
//				"                <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"                    <w:pPr>\n" + 
//				"                        <w:ind w:left=\"0\"/>\n" + 
//				"                    </w:pPr>\n" + 
//				"                    <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"                        <w:t xml:space=\"preserve\">Cell text</w:t>\n" + 
//				"                    </w:r>\n" + 
//				"                </w:p>\n" + 
//				"            </w:tc>\n" + 
//				"            <w:tc>\n" + 
//				"                <w:tcPr>\n" + 
//				"                    <w:tcW w:type=\"dxa\" w:w=\"3561\"/>\n" + 
//				"                </w:tcPr>\n" + 
//				"                <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00D15781\">\n" + 
//				"                    <w:pPr>\n" + 
//				"                        <w:ind w:left=\"0\"/>\n" + 
//				"                    </w:pPr>\n" + 
//				"                </w:p>\n" + 
//				"            </w:tc>\n" + 
//				"            <w:tc>\n" + 
//				"                <w:tcPr>\n" + 
//				"                    <w:tcW w:type=\"dxa\" w:w=\"3561\"/>\n" + 
//				"                    <w:shd w:color=\"auto\" w:fill=\"D9D9D9\" w:themeFill=\"background1\" w:themeFillShade=\"D9\" w:val=\"clear\"/>\n" + 
//				"                </w:tcPr>\n" + 
//				"                <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"                    <w:pPr>\n" + 
//				"                        <w:ind w:left=\"0\"/>\n" + 
//				"                    </w:pPr>\n" + 
//				"                    <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"                        <w:t xml:space=\"preserve\">Shaded grey</w:t>\n" + 
//				"                    </w:r>\n" + 
//				"                </w:p>\n" + 
//				"            </w:tc>\n" + 
//				"        </w:tr>\n" + 
//				"        <w:tr w:rsidR=\"00D15781\">\n" + 
//				"            <w:tc>\n" + 
//				"                <w:tcPr>\n" + 
//				"                    <w:tcW w:type=\"dxa\" w:w=\"3561\"/>\n" + 
//				"                    <w:vMerge w:val=\"restart\"/>\n" + 
//				"                </w:tcPr>\n" + 
//				"                <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"                    <w:pPr>\n" + 
//				"                        <w:ind w:left=\"0\"/>\n" + 
//				"                    </w:pPr>\n" + 
//				"                    <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"                        <w:t xml:space=\"preserve\">Vertical merge</w:t>\n" + 
//				"                    </w:r>\n" + 
//				"                </w:p>\n" + 
//				"            </w:tc>\n" + 
//				"            <w:tc>\n" + 
//				"                <w:tcPr>\n" + 
//				"                    <w:tcW w:type=\"dxa\" w:w=\"3561\"/>\n" + 
//				"                </w:tcPr>\n" + 
//				"                <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00D15781\">\n" + 
//				"                    <w:pPr>\n" + 
//				"                        <w:ind w:left=\"0\"/>\n" + 
//				"                    </w:pPr>\n" + 
//				"                </w:p>\n" + 
//				"            </w:tc>\n" + 
//				"            <w:tc>\n" + 
//				"                <w:tcPr>\n" + 
//				"                    <w:tcW w:type=\"dxa\" w:w=\"3561\"/>\n" + 
//				"                    <w:shd w:color=\"auto\" w:fill=\"D9D9D9\" w:themeFill=\"background1\" w:themeFillShade=\"D9\" w:val=\"clear\"/>\n" + 
//				"                </w:tcPr>\n" + 
//				"                <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"                    <w:pPr>\n" + 
//				"                        <w:ind w:left=\"0\"/>\n" + 
//				"                    </w:pPr>\n" + 
//				"                    <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"                        <w:t xml:space=\"preserve\">Shaded grey</w:t>\n" + 
//				"                    </w:r>\n" + 
//				"                </w:p>\n" + 
//				"            </w:tc>\n" + 
//				"        </w:tr>\n" + 
//				"        <w:tr w:rsidR=\"00D15781\">\n" + 
//				"            <w:tc>\n" + 
//				"                <w:tcPr>\n" + 
//				"                    <w:tcW w:type=\"dxa\" w:w=\"3561\"/>\n" + 
//				"                    <w:vMerge/>\n" + 
//				"                </w:tcPr>\n" + 
//				"                <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00D15781\">\n" + 
//				"                    <w:pPr>\n" + 
//				"                        <w:ind w:left=\"0\"/>\n" + 
//				"                    </w:pPr>\n" + 
//				"                </w:p>\n" + 
//				"            </w:tc>\n" + 
//				"            <w:tc>\n" + 
//				"                <w:tcPr>\n" + 
//				"                    <w:tcW w:type=\"dxa\" w:w=\"3561\"/>\n" + 
//				"                </w:tcPr>\n" + 
//				"                <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00D15781\">\n" + 
//				"                    <w:pPr>\n" + 
//				"                        <w:ind w:left=\"0\"/>\n" + 
//				"                    </w:pPr>\n" + 
//				"                </w:p>\n" + 
//				"            </w:tc>\n" + 
//				"            <w:tc>\n" + 
//				"                <w:tcPr>\n" + 
//				"                    <w:tcW w:type=\"dxa\" w:w=\"3561\"/>\n" + 
//				"                </w:tcPr>\n" + 
//				"                <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00D15781\">\n" + 
//				"                    <w:pPr>\n" + 
//				"                        <w:ind w:left=\"0\"/>\n" + 
//				"                    </w:pPr>\n" + 
//				"                </w:p>\n" + 
//				"            </w:tc>\n" + 
//				"        </w:tr>\n" + 
//				"        <w:tr w:rsidR=\"00D15781\">\n" + 
//				"            <w:tc>\n" + 
//				"                <w:tcPr>\n" + 
//				"                    <w:tcW w:type=\"dxa\" w:w=\"3561\"/>\n" + 
//				"                </w:tcPr>\n" + 
//				"                <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00D15781\">\n" + 
//				"                    <w:pPr>\n" + 
//				"                        <w:ind w:left=\"0\"/>\n" + 
//				"                    </w:pPr>\n" + 
//				"                </w:p>\n" + 
//				"            </w:tc>\n" + 
//				"            <w:tc>\n" + 
//				"                <w:tcPr>\n" + 
//				"                    <w:tcW w:type=\"dxa\" w:w=\"7122\"/>\n" + 
//				"                    <w:gridSpan w:val=\"2\"/>\n" + 
//				"                </w:tcPr>\n" + 
//				"                <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"                    <w:pPr>\n" + 
//				"                        <w:ind w:left=\"0\"/>\n" + 
//				"                    </w:pPr>\n" + 
//				"                    <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"                        <w:t xml:space=\"preserve\">Horizontal merge</w:t>\n" + 
//				"                    </w:r>\n" + 
//				"                </w:p>\n" + 
//				"            </w:tc>\n" + 
//				"        </w:tr>\n" + 
//				"    </w:tbl>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00D15781\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:ind w:left=\"0\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:ind w:left=\"0\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">(There is another document which tests tables more thoroughly)</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:pStyle w:val=\"Heading1\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">Paragraph properties</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00D15781\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:ind w:left=\"0\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">Left indent</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:jc w:val=\"center\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">Centre</w:t>\n" + 
//				"        </w:r>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">d </w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:jc w:val=\"right\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">Align Right</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:jc w:val=\"both\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">Justified text</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00D15781\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:jc w:val=\"both\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:ind w:left=\"720\"/>\n" + 
//				"            <w:jc w:val=\"both\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">Indented</w:t>\n" + 
//				"        </w:r>\n" + 
//				"        <w:del xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\" w:date=\"2009-03-11T17:57:00Z\" w:author=\"someone\" w:id=\"2\">\n" + 
//				"            <w:r>\n" + 
//				"                <w:t xml:space=\"preserve\"> indented indented indented</w:t>\n" + 
//				"            </w:r>\n" + 
//				"        </w:del>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\"> indented indented indented indented indented indented indented indented indented indented indented indented indented indented indented indented indented indented indented </w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00D15781\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:jc w:val=\"both\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:ind w:hanging=\"720\" w:left=\"1440\"/>\n" + 
//				"            <w:jc w:val=\"both\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">First line indent, Left indent, Hanging indent aaa bbb aaa bbb aaa bbb aaa bbb aaa bbb aaa bbb aaa bbb aaa bbb aaa bbb aaa bbb aaa bbb aaa bbb aaa bbb aaa bbb</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00D15781\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:jc w:val=\"both\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00D15781\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:jc w:val=\"both\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:jc w:val=\"both\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">Normal</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00665DAE\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:spacing w:after=\"400\" w:before=\"200\"/>\n" + 
//				"            <w:ind w:left=\"85\" w:right=\"85\"/>\n" + 
//				"            <w:jc w:val=\"both\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">A</w:t>\n" + 
//				"        </w:r>\n" + 
//				"        <w:ins xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\" w:date=\"2009-03-11T17:57:00Z\" w:author=\"someone\" w:id=\"3\">\n" + 
//				"            <w:r>\n" + 
//				"                <w:t xml:space=\"preserve\"> </w:t>\n" + 
//				"            </w:r>\n" + 
//				"        </w:ins>\n" + 
//				"        <w:del xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\" w:date=\"2009-03-11T17:57:00Z\" w:author=\"someone\" w:id=\"4\">\n" + 
//				"            <w:r>\n" + 
//				"                <w:t xml:space=\"preserve\"> para</w:t>\n" + 
//				"            </w:r>\n" + 
//				"        </w:del>\n" + 
//				"        <w:ins xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\" w:date=\"2009-03-11T17:57:00Z\" w:author=\"someone\" w:id=\"5\">\n" + 
//				"            <w:r>\n" + 
//				"                <w:t xml:space=\"preserve\">short </w:t>\n" + 
//				"            </w:r>\n" + 
//				"        </w:ins>\n" + 
//				"        <w:bookmarkStart w:id=\"0\" w:name=\"_GoBack\"/>\n" + 
//				"        <w:bookmarkEnd w:id=\"0\"/>\n" + 
//				"        <w:ins xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\" w:date=\"2009-03-11T17:57:00Z\" w:author=\"someone\" w:id=\"6\">\n" + 
//				"            <w:r>\n" + 
//				"                <w:t xml:space=\"preserve\">para</w:t>\n" + 
//				"            </w:r>\n" + 
//				"        </w:ins>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">graph with 10 points spacing before, 20 points after.</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:pStyle w:val=\"Heading1\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">Run properties</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00D15781\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:jc w:val=\"both\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:jc w:val=\"both\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">Font styles </w:t>\n" + 
//				"        </w:r>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:rPr>\n" + 
//				"                <w:rFonts w:ascii=\"Arial Black\" w:hAnsi=\"Arial Black\"/>\n" + 
//				"            </w:rPr>\n" + 
//				"            <w:t xml:space=\"preserve\">Aerial Black</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:jc w:val=\"both\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">Font styles </w:t>\n" + 
//				"        </w:r>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:rPr>\n" + 
//				"                <w:sz w:val=\"36\"/>\n" + 
//				"                <w:szCs w:val=\"36\"/>\n" + 
//				"            </w:rPr>\n" + 
//				"            <w:t xml:space=\"preserve\">18 point</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:jc w:val=\"both\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">Font styles </w:t>\n" + 
//				"        </w:r>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:rPr>\n" + 
//				"                <w:b/>\n" + 
//				"            </w:rPr>\n" + 
//				"            <w:t xml:space=\"preserve\">bold</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:jc w:val=\"both\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">Font styles </w:t>\n" + 
//				"        </w:r>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:rPr>\n" + 
//				"                <w:i/>\n" + 
//				"            </w:rPr>\n" + 
//				"            <w:t xml:space=\"preserve\">italic</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:jc w:val=\"both\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">Font styles </w:t>\n" + 
//				"        </w:r>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:rPr>\n" + 
//				"                <w:u w:val=\"single\"/>\n" + 
//				"            </w:rPr>\n" + 
//				"            <w:t xml:space=\"preserve\">underline</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:pStyle w:val=\"Heading1\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">Bullets &amp; numbering</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00D15781\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:jc w:val=\"both\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:jc w:val=\"both\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">Bullets</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:pStyle w:val=\"ListParagraph\"/>\n" + 
//				"            <w:numPr>\n" + 
//				"                <w:ilvl w:val=\"0\"/>\n" + 
//				"                <w:numId w:val=\"1\"/>\n" + 
//				"            </w:numPr>\n" + 
//				"            <w:jc w:val=\"both\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">Level 1</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:pStyle w:val=\"ListParagraph\"/>\n" + 
//				"            <w:numPr>\n" + 
//				"                <w:ilvl w:val=\"1\"/>\n" + 
//				"                <w:numId w:val=\"1\"/>\n" + 
//				"            </w:numPr>\n" + 
//				"            <w:jc w:val=\"both\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:lastRenderedPageBreak/>\n" + 
//				"        </w:r>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">Level 2</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00D15781\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:jc w:val=\"both\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:jc w:val=\"both\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">Numbering</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:pStyle w:val=\"ListParagraph\"/>\n" + 
//				"            <w:numPr>\n" + 
//				"                <w:ilvl w:val=\"0\"/>\n" + 
//				"                <w:numId w:val=\"2\"/>\n" + 
//				"            </w:numPr>\n" + 
//				"            <w:jc w:val=\"both\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">Level 1</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:pStyle w:val=\"ListParagraph\"/>\n" + 
//				"            <w:numPr>\n" + 
//				"                <w:ilvl w:val=\"1\"/>\n" + 
//				"                <w:numId w:val=\"2\"/>\n" + 
//				"            </w:numPr>\n" + 
//				"            <w:jc w:val=\"both\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">Level 2</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:pStyle w:val=\"ListParagraph\"/>\n" + 
//				"            <w:numPr>\n" + 
//				"                <w:ilvl w:val=\"2\"/>\n" + 
//				"                <w:numId w:val=\"2\"/>\n" + 
//				"            </w:numPr>\n" + 
//				"            <w:jc w:val=\"both\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">Level 3</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00D15781\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:ind w:left=\"720\"/>\n" + 
//				"            <w:jc w:val=\"both\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:pStyle w:val=\"Heading1\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">Images</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00D15781\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:ind w:left=\"0\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">Jpeg:</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:drawing>\n" + 
//				"                <wp:inline distB=\"0\" distL=\"0\" distR=\"0\" distT=\"0\">\n" + 
//				"                    <wp:extent cx=\"3238500\" cy=\"2362200\"/>\n" + 
//				"                    <wp:effectExtent b=\"0\" l=\"19050\" r=\"0\" t=\"0\"/>\n" + 
//				"                    <wp:docPr descr=\"C:\\Documents and Settings\\Jason Harrop\\My Documents\\tmp-test-docs\\pangolin.jpeg\" id=\"1\" name=\"Picture 1\"/>\n" + 
//				"                    <wp:cNvGraphicFramePr>\n" + 
//				"                        <a:graphicFrameLocks noChangeAspect=\"true\"/>\n" + 
//				"                    </wp:cNvGraphicFramePr>\n" + 
//				"                    <a:graphic>\n" + 
//				"                        <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">\n" + 
//				"                            <pic:pic>\n" + 
//				"                                <pic:nvPicPr>\n" + 
//				"                                    <pic:cNvPr descr=\"C:\\Documents and Settings\\Jason Harrop\\My Documents\\tmp-test-docs\\pangolin.jpeg\" id=\"0\" name=\"Picture 1\"/>\n" + 
//				"                                    <pic:cNvPicPr>\n" + 
//				"                                        <a:picLocks noChangeArrowheads=\"true\" noChangeAspect=\"true\"/>\n" + 
//				"                                    </pic:cNvPicPr>\n" + 
//				"                                </pic:nvPicPr>\n" + 
//				"                                <pic:blipFill>\n" + 
//				"                                    <a:blip r:embed=\"rId8Lblagh\"/>\n" + 
//				"                                    <a:srcRect/>\n" + 
//				"                                    <a:stretch>\n" + 
//				"                                        <a:fillRect/>\n" + 
//				"                                    </a:stretch>\n" + 
//				"                                </pic:blipFill>\n" + 
//				"                                <pic:spPr bwMode=\"auto\">\n" + 
//				"                                    <a:xfrm>\n" + 
//				"                                        <a:off x=\"0\" y=\"0\"/>\n" + 
//				"                                        <a:ext cx=\"3238500\" cy=\"2362200\"/>\n" + 
//				"                                    </a:xfrm>\n" + 
//				"                                    <a:prstGeom prst=\"rect\">\n" + 
//				"                                        <a:avLst/>\n" + 
//				"                                    </a:prstGeom>\n" + 
//				"                                    <a:noFill/>\n" + 
//				"                                    <a:ln w=\"9525\">\n" + 
//				"                                        <a:noFill/>\n" + 
//				"                                        <a:miter lim=\"800000\"/>\n" + 
//				"                                        <a:headEnd/>\n" + 
//				"                                        <a:tailEnd/>\n" + 
//				"                                    </a:ln>\n" + 
//				"                                </pic:spPr>\n" + 
//				"                            </pic:pic>\n" + 
//				"                        </a:graphicData>\n" + 
//				"                    </a:graphic>\n" + 
//				"                </wp:inline>\n" + 
//				"            </w:drawing>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00D15781\"/>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00D15781\"/>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">Gif (scaled):</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:lastRenderedPageBreak/>\n" + 
//				"        </w:r>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:drawing>\n" + 
//				"                <wp:inline distB=\"0\" distL=\"0\" distR=\"0\" distT=\"0\">\n" + 
//				"                    <wp:extent cx=\"2809875\" cy=\"5473022\"/>\n" + 
//				"                    <wp:effectExtent b=\"0\" l=\"19050\" r=\"9525\" t=\"0\"/>\n" + 
//				"                    <wp:docPr descr=\"Escher: Liberation\" id=\"2\" name=\"Picture 2\"/>\n" + 
//				"                    <wp:cNvGraphicFramePr>\n" + 
//				"                        <a:graphicFrameLocks noChangeAspect=\"true\"/>\n" + 
//				"                    </wp:cNvGraphicFramePr>\n" + 
//				"                    <a:graphic>\n" + 
//				"                        <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">\n" + 
//				"                            <pic:pic>\n" + 
//				"                                <pic:nvPicPr>\n" + 
//				"                                    <pic:cNvPr descr=\"Escher: Liberation\" id=\"0\" name=\"Picture 2\"/>\n" + 
//				"                                    <pic:cNvPicPr>\n" + 
//				"                                        <a:picLocks noChangeArrowheads=\"true\" noChangeAspect=\"true\"/>\n" + 
//				"                                    </pic:cNvPicPr>\n" + 
//				"                                </pic:nvPicPr>\n" + 
//				"                                <pic:blipFill>\n" + 
//				"                                    <a:blip r:embed=\"rId9Lblagh\"/>\n" + 
//				"                                    <a:srcRect/>\n" + 
//				"                                    <a:stretch>\n" + 
//				"                                        <a:fillRect/>\n" + 
//				"                                    </a:stretch>\n" + 
//				"                                </pic:blipFill>\n" + 
//				"                                <pic:spPr bwMode=\"auto\">\n" + 
//				"                                    <a:xfrm>\n" + 
//				"                                        <a:off x=\"0\" y=\"0\"/>\n" + 
//				"                                        <a:ext cx=\"2810209\" cy=\"5473672\"/>\n" + 
//				"                                    </a:xfrm>\n" + 
//				"                                    <a:prstGeom prst=\"rect\">\n" + 
//				"                                        <a:avLst/>\n" + 
//				"                                    </a:prstGeom>\n" + 
//				"                                    <a:noFill/>\n" + 
//				"                                    <a:ln w=\"9525\">\n" + 
//				"                                        <a:noFill/>\n" + 
//				"                                        <a:miter lim=\"800000\"/>\n" + 
//				"                                        <a:headEnd/>\n" + 
//				"                                        <a:tailEnd/>\n" + 
//				"                                    </a:ln>\n" + 
//				"                                </pic:spPr>\n" + 
//				"                            </pic:pic>\n" + 
//				"                        </a:graphicData>\n" + 
//				"                    </a:graphic>\n" + 
//				"                </wp:inline>\n" + 
//				"            </w:drawing>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00D15781\"/>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00D15781\"/>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">Png (from </w:t>\n" + 
//				"        </w:r>\n" + 
//				"        <w:hyperlink xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\" r:id=\"rId10\" w:history=\"true\">\n" + 
//				"            <w:r>\n" + 
//				"                <w:rPr>\n" + 
//				"                    <w:rStyle w:val=\"Hyperlink\"/>\n" + 
//				"                </w:rPr>\n" + 
//				"                <w:t xml:space=\"preserve\">http://davidpritchard.org/images/pacsoc-s1b.png</w:t>\n" + 
//				"            </w:r>\n" + 
//				"        </w:hyperlink>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\"> )</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:drawing>\n" + 
//				"                <wp:inline distB=\"0\" distL=\"0\" distR=\"0\" distT=\"0\">\n" + 
//				"                    <wp:extent cx=\"4286250\" cy=\"3343275\"/>\n" + 
//				"                    <wp:effectExtent b=\"0\" l=\"19050\" r=\"0\" t=\"0\"/>\n" + 
//				"                    <wp:docPr descr=\"http://davidpritchard.org/images/pacsoc-s1b.png\" id=\"5\" name=\"Picture 5\"/>\n" + 
//				"                    <wp:cNvGraphicFramePr>\n" + 
//				"                        <a:graphicFrameLocks noChangeAspect=\"true\"/>\n" + 
//				"                    </wp:cNvGraphicFramePr>\n" + 
//				"                    <a:graphic>\n" + 
//				"                        <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">\n" + 
//				"                            <pic:pic>\n" + 
//				"                                <pic:nvPicPr>\n" + 
//				"                                    <pic:cNvPr descr=\"http://davidpritchard.org/images/pacsoc-s1b.png\" id=\"0\" name=\"Picture 5\"/>\n" + 
//				"                                    <pic:cNvPicPr>\n" + 
//				"                                        <a:picLocks noChangeArrowheads=\"true\" noChangeAspect=\"true\"/>\n" + 
//				"                                    </pic:cNvPicPr>\n" + 
//				"                                </pic:nvPicPr>\n" + 
//				"                                <pic:blipFill>\n" + 
//				"                                    <a:blip r:embed=\"rId11Lblagh\"/>\n" + 
//				"                                    <a:srcRect/>\n" + 
//				"                                    <a:stretch>\n" + 
//				"                                        <a:fillRect/>\n" + 
//				"                                    </a:stretch>\n" + 
//				"                                </pic:blipFill>\n" + 
//				"                                <pic:spPr bwMode=\"auto\">\n" + 
//				"                                    <a:xfrm>\n" + 
//				"                                        <a:off x=\"0\" y=\"0\"/>\n" + 
//				"                                        <a:ext cx=\"4286250\" cy=\"3343275\"/>\n" + 
//				"                                    </a:xfrm>\n" + 
//				"                                    <a:prstGeom prst=\"rect\">\n" + 
//				"                                        <a:avLst/>\n" + 
//				"                                    </a:prstGeom>\n" + 
//				"                                    <a:noFill/>\n" + 
//				"                                    <a:ln w=\"9525\">\n" + 
//				"                                        <a:noFill/>\n" + 
//				"                                        <a:miter lim=\"800000\"/>\n" + 
//				"                                        <a:headEnd/>\n" + 
//				"                                        <a:tailEnd/>\n" + 
//				"                                    </a:ln>\n" + 
//				"                                </pic:spPr>\n" + 
//				"                            </pic:pic>\n" + 
//				"                        </a:graphicData>\n" + 
//				"                    </a:graphic>\n" + 
//				"                </wp:inline>\n" + 
//				"            </w:drawing>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00D15781\"/>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:ind w:left=\"0\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">(TODO: we really should have both 2003 &amp; 2007 pictures)</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00D15781\"/>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:br w:type=\"page\"/>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:ind w:left=\"0\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:lastRenderedPageBreak/>\n" + 
//				"        </w:r>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">That was a page break</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00D15781\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:ind w:left=\"0\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:ind w:left=\"0\"/>\n" + 
//				"            <w:jc w:val=\"both\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">Here is some change tracking. </w:t>\n" + 
//				"        </w:r>\n" + 
//				"        <w:ins w:author=\"Jason Harrop\" w:date=\"2007-12-09T10:14:00Z\" w:id=\"1\">\n" + 
//				"            <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"                <w:t xml:space=\"preserve\">An insertion</w:t>\n" + 
//				"            </w:r>\n" + 
//				"        </w:ins>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\"> Followed by</w:t>\n" + 
//				"        </w:r>\n" + 
//				"        <w:del w:author=\"Jason Harrop\" w:date=\"2007-12-09T10:14:00Z\" w:id=\"2\">\n" + 
//				"            <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"                <w:delText xml:space=\"preserve\">\n" + 
//				"                    <w:r>\n" + 
//				"                        <w:t xml:space=\"preserve\"> A deletion</w:t>\n" + 
//				"                    </w:r>\n" + 
//				"                </w:delText>\n" + 
//				"            </w:r>\n" + 
//				"        </w:del>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">.</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00D15781\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:ind w:left=\"0\"/>\n" + 
//				"            <w:jc w:val=\"both\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00945132\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:ind w:left=\"0\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">This line contains a soft return</w:t>\n" + 
//				"        </w:r>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:br/>\n" + 
//				"        </w:r>\n" + 
//				"        <w:r xmlns:xalan=\"http://xml.apache.org/xalan\" xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">\n" + 
//				"            <w:t xml:space=\"preserve\">and here it continues</w:t>\n" + 
//				"        </w:r>\n" + 
//				"    </w:p>\n" + 
//				"    <w:p w:rsidR=\"00D15781\" w:rsidRDefault=\"00D15781\">\n" + 
//				"        <w:pPr>\n" + 
//				"            <w:ind w:left=\"0\"/>\n" + 
//				"        </w:pPr>\n" + 
//				"    </w:p>\n" + 
//				"    <w:sectPr w:rsidR=\"00D15781\">\n" + 
//				"        <w:pgSz w:code=\"9\" w:h=\"16839\" w:w=\"11907\"/>\n" + 
//				"        <w:pgMar w:bottom=\"720\" w:footer=\"720\" w:gutter=\"0\" w:header=\"720\" w:left=\"720\" w:right=\"720\" w:top=\"720\"/>\n" + 
//				"        <w:cols w:space=\"720\"/>\n" + 
//				"        <w:docGrid w:linePitch=\"360\"/>\n" + 
//				"    </w:sectPr>\n" + 
//				"</w:body>";
		
		Body newBody = (Body) XmlUtils.unwrap(
				XmlUtils.unmarshalString(contentStr));
		
		// 4. Display the result as a PDF
		// To do this, we'll replace the body in the newer document
		((Document)newerPackage.getMainDocumentPart().getJaxbElement()).setBody(newBody);

		if (DIVIDE_AND_CONQUER) {
			// No image support at present
		} else {
			handleRels(pd, newerPackage.getMainDocumentPart()); // TODO: that needs work, for more complex input
		}
		
		
		if (DOCX_SAVE) {
			newerPackage.save(new File("/Users/faresyoussef/Desktop/testing differences/OUT_CompareDocuments.docx"));
		}
		
		if (PDF_SAVE) {
			
			newerPackage.setFontMapper(new IdentityPlusMapper());	
			
			boolean saveFO = false;
			String outputfilepath = System.getProperty("user.dir") +  "/OUT_CompareDocuments.pdf";
			
			// FO exporter setup (required)
			// .. the FOSettings object
	    	FOSettings foSettings = Docx4J.createFOSettings();
			if (saveFO) {
				foSettings.setFoDumpFile(new java.io.File(System.getProperty("user.dir") +  "/OUT_CompareDocuments..fo"));
			}
			foSettings.setOpcPackage(newerPackage);
			// Document format: 
			// The default implementation of the FORenderer that uses Apache Fop will output
			// a PDF document if nothing is passed via 
			// foSettings.setApacheFopMime(apacheFopMime)
			// apacheFopMime can be any of the output formats defined in org.apache.fop.apps.MimeConstants or
			// FOSettings.INTERNAL_FO_MIME if you want the fo document as the result.
			
			
			// exporter writes to an OutputStream.		
			OutputStream os = new java.io.FileOutputStream(outputfilepath);
	    	
	
			//Don't care what type of exporter you use
			Docx4J.toFO(foSettings, os, Docx4J.FLAG_NONE);
			
			System.out.println("Saved " + System.getProperty("user.dir")  +  "/OUT_CompareDocuments.pdf");
		}
				
	}

	/**
		 In the general case, you need to handle relationships.
		 Although not necessary in this simple example, 
		 we do it anyway for the purposes of illustration.
	 * @throws InvalidFormatException 
		 
	 */
	private static void handleRels(Differencer pd, MainDocumentPart newMDP) throws InvalidFormatException {
		
		RelationshipsPart rp = newMDP.getRelationshipsPart(); 
		System.out.println("before: \n" + rp.getXML());		
		
		// Since we are going to add rels appropriate to the docs being 
		// compared, for neatness and to avoid duplication
		// (duplication of internal part names is fatal in Word,
		//  and export xslt makes images internal, though it does avoid duplicating
		//  a part ), 
		// remove any existing rels which point to images
		List<Relationship> relsToRemove = new ArrayList<Relationship>();
		for (Relationship r : rp.getRelationships().getRelationship() ) {
			//  Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
			if (r.getType().equals(Namespaces.IMAGE)) {
				relsToRemove.add(r);
			}
		}						
		for (Relationship r : relsToRemove) {				
			rp.removeRelationship(r);
		}
		
		// Now add the rels we composed
		Map<Relationship, Part> newRels = pd.getComposedRels();
		for (Relationship nr : newRels.keySet()) {	
			
			if (nr.getTargetMode()!=null 
					&& nr.getTargetMode().equals("External")) {
				
				newMDP.getRelationshipsPart().getRelationships().getRelationship().add(nr);
				
			} else {
				
				Part part = newRels.get(nr);
				if (part==null) {
					System.out.println("ERROR! Couldn't find part for rel " + nr.getId() + "  " + nr.getTargetMode() );
				} else {
					
					if (part instanceof BinaryPart) { // ensure contents are loaded, before moving to new pkg
						((BinaryPart)part).getBuffer();
					}
					
					newMDP.addTargetPart(part, AddPartBehaviour.RENAME_IF_NAME_EXISTS, nr.getId());
				}
			}
		}
		
		System.out.println("after: \n" + rp.getXML());
		
	}
	
	
		

}
