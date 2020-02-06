package org.obin.docgen;

import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import com.sun.javadoc.ClassDoc;
import com.sun.javadoc.MethodDoc;
import com.sun.javadoc.Parameter;
import com.sun.javadoc.RootDoc;

public class JavaDocReader {
	private static RootDoc root;

	public static class CustomDoclet {
		public static boolean start(RootDoc root) {
			JavaDocReader.root = root;
			return true;
		}
	}

	public static void process(List<String> sourcePaths, List<String> javapackages, List<String> excludePackages,
			String outputDir) throws Exception {

		String paths = sourcePaths.stream().collect(Collectors.joining(";"));
		String includes = javapackages.stream().collect(Collectors.joining(":"));
		String excludes = excludePackages.stream().collect(Collectors.joining(":"));

		List<String> argsOrderList = new ArrayList<>();
		argsOrderList.add("-doclet");
		argsOrderList.add(CustomDoclet.class.getName());

		if (paths != null && paths.length() > 0) {
			argsOrderList.add("-sourcepath");
			argsOrderList.add(paths);
		}

		argsOrderList.add("-encoding");
		argsOrderList.add("utf-8");
		argsOrderList.add("-verbose");

		if (includes != null && includes.length() > 0) {
			argsOrderList.add("-subpackages");
			argsOrderList.add(includes);
		}

		if (excludes != null && excludes.length() > 0) {
			argsOrderList.add("-exclude");
			argsOrderList.add(excludes);
		}

		String[] args = argsOrderList.toArray(new String[argsOrderList.size()]);
		System.out.println(Arrays.toString(args));
		com.sun.tools.javadoc.Main.execute(args);

		File file = new File(outputDir);

		List<List<String>> allData = new ArrayList<>();
		ClassDoc[] classes = root.classes();

		allData.add(Arrays.asList("方法", "参数", "备注"));
		if (classes != null) {
			for (int i = 0; i < classes.length; ++i) {
				if (classes[i].containingClass() == null && classes[i].isPublic()) {

					String filename = classes[i].qualifiedTypeName();
					MethodDoc[] methodocs = classes[i].methods();

					for (MethodDoc me : methodocs) {
						if (me.name().startsWith("set")) {
							continue;
						}

						List<String> info = new ArrayList<>();
						info.add(filename + "." + me.name());
						info.add(getFullparameter(me.parameters()));
						info.add(me.commentText());

						allData.add(info);
					}

				}
			}
		}
		root = null;

		ExcelUtil eu = new ExcelUtil(outputDir);
		eu.writeExcel(allData);
	}

	public static String getFullparameter(Parameter[] para) {
		StringBuffer sb = new StringBuffer();
		Stream.of(para).forEach(a -> sb.append(a.typeName() + " " + a.name() + "  "));
		return sb.toString();
	}
}
