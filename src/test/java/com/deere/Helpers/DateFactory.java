
/* 
 * Project    : DealerPath
 * Script     : DateFactory
 * Author     : Shrishail Baddi
 * Date       : April.03.2018
 * Last Modified On:
 * Modified By :
 */

package com.deere.Helpers;

import java.text.SimpleDateFormat;
import java.util.Date;

public class DateFactory {

	public static String formate(Date date) {
		if (date == null)
			return "";
		return String.format("%1$tY%1$tm%1$td%1$tH%1$tM%1$tS%1$tL", date);

	}

	/**
	 * 2010-09-10 21:08:17
	 */
	public static String formateYMDHMS(Date date) {
		if (date == null) {
			return "";
		}
		return String.format("%1$tY-%1$tm-%1$td %1$tH:%1$tM:%1$tS", date);
	}

	/**
	 * 2010-09-10
	 */
	public static String formateYMD(Date date) {
		if (date == null) {
			return "";
		}
		return String.format("%1$tY-%1$tm-%1$td", date);
	}

	public static String captureYear(Date date) {
		if (date == null) {
			return "";
		}
		return DateFactory.formateYMD(date).split("-")[0];
	}

	/**
	 * :09-10
	 */
	public static String formateMD(Date date) {
		if (date == null) {
			return "";
		}
		return String.format("%1$tMM-%1$td", date);
	}


	public static boolean isThisDateValid(String dateToValidate, String dateFromat) {

		if (dateToValidate == null) {
			return false;
		}

		SimpleDateFormat sdf = new SimpleDateFormat(dateFromat);
		sdf.setLenient(false);

		try {

			// if not valid, it will throw ParseException
			Date date = sdf.parse(dateToValidate);
			 System.out.println(date);

		} catch (Exception e) {

			e.printStackTrace();
			return false;
		}

		return true;
	}

}
