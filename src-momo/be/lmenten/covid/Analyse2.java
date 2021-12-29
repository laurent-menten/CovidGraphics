package be.lmenten.covid;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.ObjectOutputStream;
import java.time.LocalDate;
import java.time.Month;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.util.logging.Level;

import javax.imageio.ImageIO;
import javax.swing.filechooser.FileSystemView;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.image.BufferedImage;

import com.sun.star.frame.XComponentLoader;
import com.sun.star.sheet.XCellRangesQuery;
import com.sun.star.sheet.XSheetCellRanges;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.table.CellRangeAddress;
import com.sun.star.table.XCellRange;

import be.lmenten.utils.logging.AnsiLogFormatter;
import be.lmenten.utils.logging.LogRegExPackageFilter;
import lo.utils.Calc;
import lo.utils.Lo;

public class Analyse2
{
	private static final LocalDate FIRST_DAY = LocalDate.of( 2020, Month.MARCH, 1 );
	private static final LocalDate LAST_DAY = LocalDate.now();
	private static final int DAYS_COUNT = (int) ChronoUnit.DAYS.between( FIRST_DAY, LAST_DAY );

	private static final LocalDate FIRST_DATASET = LocalDate.of( 2020, Month.MARCH, 1 );
	private static final LocalDate LAST_DATASET = LocalDate.now();
	private static final int DATASETS_COUNT = (int) ChronoUnit.DAYS.between( FIRST_DATASET, LAST_DATASET );

	// ------------------------------------------------------------------------

	static final DateTimeFormatter FILENAME_DATE_FORMAT
		= DateTimeFormatter.ofPattern( "yyyyMMdd" );

	static final String FILENAME_PATTERN
		= FileSystemView.getFileSystemView().getDefaultDirectory().getPath()
			+ "/Covid19BE/data/COVID19BE_%s.xlsx";

	// ========================================================================
	// =
	// ========================================================================

	public static void main( String[] args )
	{
		AnsiLogFormatter.install();
		LogRegExPackageFilter.install( "be.lmenten.*", Level.ALL );

		XComponentLoader loader = Lo.loadOffice();

		// --------------------------------------------------------------------
		// - Process data sets ------------------------------------------------
		// --------------------------------------------------------------------

		Integer [][] data = new Integer [DAYS_COUNT][DATASETS_COUNT];

		for( LocalDate dataset = FIRST_DAY ; dataset.isBefore(LAST_DAY) ; dataset = dataset.plusDays( 1l ) )
		{
			int datasetIndex = (int) ChronoUnit.DAYS.between( FIRST_DATASET, dataset );

			String filename = String.format(  FILENAME_PATTERN, dataset.format(FILENAME_DATE_FORMAT ) );

			System.out.print( datasetIndex + ": " + filename );

			File file = new File( filename );
			if( ! file.exists() )
			{
				System.out.println( " - File not found." );
				continue;
			}

			// ----------------------------------------------------------------

			XSpreadsheetDocument doc = Calc.openDoc( filename, loader );
			if( doc == null )
			{
				System.out.println( " - Failed to open document." );
				continue;
			}

			XSpreadsheet sheet = Calc.getSheet( doc, "MORT" );
			if( sheet == null )
			{
				System.out.println( " - Failed to get 'MORT' sheet." );
				Lo.closeDoc( doc );
				continue;
			}

			XCellRange cellRange = Calc.findUsedRange( sheet );
			XCellRangesQuery query = Lo.qi( XCellRangesQuery.class, cellRange );
			XSheetCellRanges ranges = query.queryVisibleCells();
			CellRangeAddress [] addresses = ranges.getRangeAddresses();

//			final int colCount = addresses[0].EndColumn - addresses[0].StartColumn + 1;
			final int rowCount = (addresses[0].EndRow - addresses[0].StartRow);
//			final String colNames = LoSheetHelper.getColNamesList( sheet, addresses[0].StartRow, colCount );

			System.out.print( " (" + rowCount + " rows)"  );
			System.out.flush();

			// ----------------------------------------------------------------

			// 0    1      2        3   4
		    // DATE REGION AGEGROUP SEX DEATHS

			long chronoStart = System.currentTimeMillis();
			
			for( int i = addresses[0].StartRow + 1 ; i <= addresses[0].EndRow ; i++ )
			{
				LocalDate day = LoSheetHelper.getDate( sheet, 0, i );
				int dayIndex = (int) ChronoUnit.DAYS.between( FIRST_DAY, day );

//				String region = (String) Calc.getString( sheet, 1, i );

//				String agegroupRaw = (String) Calc.getString( sheet, 2, i );
//				AgeGroup25 agegroup = AgeGroup25.lookup( agegroupRaw );

//				String sex = (String) Calc.getString( sheet, 3, i );

				int deaths = LoSheetHelper.getInteger( sheet, 4, i );

				// ------------------------------------------------------------

				// FILTER here ...

				// ------------------------------------------------------------

				if( data[dayIndex][datasetIndex] == null )
				{
					data[dayIndex][datasetIndex] = Integer.valueOf( 0 );
				}

				data[dayIndex][datasetIndex] += deaths;
			}

			long chrono = System.currentTimeMillis() - chronoStart;
			System.out.printf( " : %,d ms.\n", chrono );

			// ----------------------------------------------------------------

			Lo.closeDoc( doc );
		}

		// --------------------------------------------------------------------
		// - Save data --------------------------------------------------------
		// --------------------------------------------------------------------

		try( FileOutputStream fos = new FileOutputStream( "data-decès-covid-all.bin" ) )
		{
			try( ObjectOutputStream oos = new ObjectOutputStream( fos ) )
			{
				oos.writeObject( FIRST_DAY );
				oos.writeObject( LAST_DAY );
				oos.writeObject( FIRST_DATASET );
				oos.writeObject( LAST_DATASET );
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
		
		// --------------------------------------------------------------------
		// - Create image -----------------------------------------------------
		// --------------------------------------------------------------------

		BufferedImage bufferedImage = new BufferedImage( DAYS_COUNT, DATASETS_COUNT, BufferedImage.TYPE_INT_RGB );
		Graphics2D g2d = bufferedImage.createGraphics();

		g2d.setColor( Color.LIGHT_GRAY );
		g2d.fillRect( 0, 0, DAYS_COUNT, DATASETS_COUNT );

		for( int day = 0; day < DAYS_COUNT ; day++ )
		{
			Integer lastValue = data[day][0];

			for( int dataset = 0 ; dataset < DAYS_COUNT ; dataset++ )
			{
				Integer currentValue = data[day][dataset]; 
				
				if( currentValue != null )
				{
					bufferedImage.setRGB( dataset, day, 0xFFFFFF );

					if( lastValue != null )
					{
						if( currentValue < lastValue )
						{
							bufferedImage.setRGB( dataset, day, 0x00FF7F );
						}
						else if( currentValue > lastValue )
						{
							bufferedImage.setRGB( dataset, day, 0xFF0000 );
						}
					}
					else
					{
						bufferedImage.setRGB( dataset, day, 0x00FF00 );
					}

					lastValue = currentValue;
				}
			}
		}

		g2d.dispose();

		try
		{
			File outputfile = new File( "image.jpg" );
			ImageIO.write( bufferedImage, "jpg", outputfile );
		}
		catch( IOException ex )
		{
			ex.printStackTrace();
		}

		// --------------------------------------------------------------------
		// - Terminate --------------------------------------------------------
		// --------------------------------------------------------------------

		Lo.closeOffice();
	}
}
