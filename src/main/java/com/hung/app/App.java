package com.hung.app;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

import org.docx4j.TraversalUtil;
import org.docx4j.finders.SectPrFinder;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.SectPr;

/**
 * Hello world!
 *
 */
public class App
{
	


	public static void main( String[] args )
    {
        System.out.println("Nhập đường dẫn folder: ");
        String pathFile = new Scanner(System.in).nextLine();
        pathFile = pathFile.trim();
        File folder = new File(pathFile);
        for(File file : folder.listFiles()) {
            try {
    			removeHFFromFile(file);
    		} catch (Exception e) {
    			// TODO Auto-generated catch block
    			e.printStackTrace();
    		}
        }
    }
    
    public static void removeHFFromFile(File f) throws Exception {

		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(f);

		MainDocumentPart mdp = wordMLPackage.getMainDocumentPart();

		// Remove from sectPr
		SectPrFinder finder = new SectPrFinder(mdp);
		new TraversalUtil(mdp.getContent(), finder);
		for (SectPr sectPr : finder.getSectPrList()) {
			sectPr.getEGHdrFtrReferences().clear();
		}

		// Remove rels
		List<Relationship> hfRels = new ArrayList<Relationship>();
		for (Relationship rel : mdp.getRelationshipsPart().getRelationships().getRelationship()) {

			if (rel.getType().equals(Namespaces.HEADER) || rel.getType().equals(Namespaces.FOOTER)) {
				hfRels.add(rel);
			}
		}
		for (Relationship rel : hfRels) {
			mdp.getRelationshipsPart().removeRelationship(rel);
		}

		wordMLPackage.save(f);
	}
}
