package com.file.msexcel.util;

import java.util.Arrays;
import java.util.List;

public class ObservacionesSplitter {

	public static List<String> split(String cellValueString, String breakChar1) {
		
		String[] result = cellValueString.split(breakChar1);
		return Arrays.asList(result);

	}

}
