package org.obin.docgen;

import java.util.Arrays;

import javax.swing.event.ListSelectionEvent;

public class DocGen {

	public static void main(String[] args) {
	 
		try {
			JavaDocReader.process(Arrays.asList("E:/hy67/jjgpp/jjgpp/custom/jjgppcore/src" ),
					Arrays.asList("jjgpp"),Arrays.asList("jjgpp.hybris.core.jalo","jjgpp.hybris.core.oa","jjgpp.hybris.core.forms","jjgpp.hybris.core.oa2","jjgpp.hybris.core.oa3","jjgpp.hybris.core.chinapay.dto","jjgpp.hybris.core.dto"),
					"E:\\hy67\\jjgpp\\jjgpp\\custom\\jjgppcore\\all.xlsx");
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}