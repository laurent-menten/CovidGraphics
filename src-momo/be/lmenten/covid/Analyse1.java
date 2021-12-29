package be.lmenten.covid;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.ObjectOutputStream;
import java.time.LocalDate;
import java.util.Map;
import java.util.TreeMap;

import com.sun.star.frame.XComponentLoader;
import com.sun.star.sheet.XCellRangesQuery;
import com.sun.star.sheet.XSheetCellRanges;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.table.CellRangeAddress;
import com.sun.star.table.XCellRange;

import lo.utils.Calc;
import lo.utils.Lo;

//
// https://statbel.fgov.be/en/open-data/number-deaths-day-sex-district-age
// https://statbel.fgov.be/sites/default/files/files/opendata/deathday/DEMO_DEATH_OPEN.xlsx
//

public class Analyse1
{
	public static void main( String[] args )
		throws FileNotFoundException, IOException
	{
		XComponentLoader loader = Lo.loadOffice();

		XSpreadsheetDocument doc = Calc.openDoc( "DEMO_DEATH_OPEN.xlsx", loader );
//	    XDocumentPropertiesSupplier dps = Lo.qi( XDocumentPropertiesSupplier.class, doc );
//	    XDocumentProperties dp = dps.getDocumentProperties();

	    XSpreadsheet sheet = Calc.getSheet( doc, 0 );

		XCellRange cellRange = Calc.findUsedRange( sheet );
		XCellRangesQuery query = Lo.qi( XCellRangesQuery.class, cellRange );
		XSheetCellRanges ranges = query.queryVisibleCells();
		CellRangeAddress [] addresses = ranges.getRangeAddresses();

		final int colCount = addresses[0].EndColumn - addresses[0].StartColumn + 1;
		final int rowCount = (addresses[0].EndRow - addresses[0].StartRow);
		final String colNames = LoSheetHelper.getColNamesList( sheet, addresses[0].StartRow, colCount );

		System.out.println( "columns: " + colCount );
		System.out.println( "names: " + colNames );
		System.out.println( "rows: " + rowCount );

		// ====================================================================
		// =
		// ====================================================================

		//	AGE
		//		DAY
		//			YEARS []

		Map<AgeGroup25,Map<String,Integer[]>> data = new TreeMap<>();

		long start = System.currentTimeMillis();
		for( int i = addresses[0].StartRow + 1 ; i <= addresses[0].EndRow ; i++ )
		{
			if( (i % 1000) == 0 )
			{
				long duration = System.currentTimeMillis() - start;
				
				double p = ((i *100.d) / rowCount);
				long eta = (long) ((((double)duration / p) * (100.d-p)) / 1000.d);
				long m = eta / 60l;
				long s = eta % 60l;

				System.out.printf( "%2.2f %% (e.t.a. %d'%d) ...\n", p, m, s );
			}

			// 0      1       2        3      4           5       6       7       8
			// CD_ARR,CD_PROV,CD_REGIO,CD_SEX,CD_AGEGROUP,DT_DATE,NR_YEAR,NR_WEEK,MS_NUM_DEATH

			// ----------------------------------------------------------------

			String agegroup = (String) Calc.getString( sheet, 4, i );
			AgeGroup25 ageClass = AgeGroup25.lookup( agegroup );

			Map<String,Integer[]> dayDataMap = data.get( ageClass );
			if( dayDataMap == null )
			{
				dayDataMap = new TreeMap<>();
				data.put( ageClass, dayDataMap );
			}

			// ----------------------------------------------------------------

			LocalDate date = LoSheetHelper.getDate( sheet, 5, i );
			String day = String.format( "%02d.%02d", date.getMonthValue(), date.getDayOfMonth() );
	
			Integer [] dayData = dayDataMap.get( day );
			if( dayData == null )
			{
				dayData = new Integer [ (2021 - 2009) + 1 ]; 
				dayDataMap.put( day, dayData );
			}
	
			// ----------------------------------------------------------------

			int yearIndex = date.getYear() - 2009;
			int deaths = LoSheetHelper.getInteger( sheet, 8, i );

			if( dayData[yearIndex] == null )
			{
				dayData[yearIndex] = deaths;
			}
			else
			{
				dayData[yearIndex] += deaths;
			}
		}

		// ====================================================================
		// =
		// ====================================================================

		data.forEach(
			(key, value) -> 
			{
				System.out.printf( "%s\n", key );
				System.out.printf( "Day, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018, 2019, 2020, 2021 \n" );

				value.forEach(
					(key2, value2) ->
					{
						System.out.printf( "%s, ", key2 );

						for( int i = 0 ; i < value2.length ; i++ )
						{
							if( value2[i] != null )
							{
								System.out.printf( "%d, ", value2[i] );
							}
							else if( key2.equals( "02.29" ) )
							{
								System.out.printf( ", "); // no data
							}
							else
							{
								System.out.printf( "0, "); // no deaths
							}
						}

						System.out.println();
					}
				);
			}
		);

		// --------------------------------------------------------------------
		// - Save data --------------------------------------------------------
		// --------------------------------------------------------------------

		try( FileOutputStream fos = new FileOutputStream( "data-mortalité.bin" ) )
		{
			try( ObjectOutputStream oos = new ObjectOutputStream( fos ) )
			{
				oos.writeObject( data );
			}
		}
		catch( FileNotFoundException ex )
		{
			ex.printStackTrace();
		}
		catch( IOException ex )
		{
			ex.printStackTrace();
		}
		
		// ====================================================================
		// =
		// ====================================================================

	    Lo.closeDoc( doc );
	    Lo.closeOffice();
	}			
}
