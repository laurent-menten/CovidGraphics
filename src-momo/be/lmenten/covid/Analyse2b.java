package be.lmenten.covid;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.time.LocalDate;
import java.time.temporal.ChronoUnit;

import javax.imageio.ImageIO;

public class Analyse2b {

	public static void main( String[] args )
	{
		// --------------------------------------------------------------------
		// - Load data --------------------------------------------------------
		// --------------------------------------------------------------------

		LocalDate FIRST_DAY = null;
		LocalDate LAST_DAY = null;

		LocalDate FIRST_DATASET = null;
		LocalDate LAST_DATASET = null;

		Integer [][] data = null;

		try( FileInputStream fis = new FileInputStream( "data-decès-covid-all.bin" ) )
		{
			try( ObjectInputStream ois = new ObjectInputStream( fis ) )
			{
				FIRST_DAY = (LocalDate) ois.readObject();
				LAST_DAY = (LocalDate) ois.readObject();

				FIRST_DATASET = (LocalDate) ois.readObject();
				LAST_DATASET = (LocalDate) ois.readObject();

				data = (Integer [][]) ois.readObject();
			}
		}
		catch( ClassNotFoundException ex )
		{
			ex.printStackTrace();
			System.exit( -1 );
		}
		catch( FileNotFoundException ex )
		{
			ex.printStackTrace();
			System.exit( -1 );
		}
		catch( IOException ex )
		{
			ex.printStackTrace();
			System.exit( -1 );
		}

		int DAYS_COUNT = (int) ChronoUnit.DAYS.between( FIRST_DAY, LAST_DAY );
		int DATASETS_COUNT = (int) ChronoUnit.DAYS.between( FIRST_DATASET, LAST_DATASET );

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
			File outputfile = new File( "image.bmp" );
			ImageIO.write( bufferedImage, "bmp", outputfile );
		}
		catch( IOException ex )
		{
			ex.printStackTrace();
		}

	}
}
