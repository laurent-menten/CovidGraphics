package be.lmenten.covid.data;

import java.util.Map;

import be.lmenten.covid.AgeGroup25;

public class StatDay
{
	private int month;
	private int day;

	private StatDayDataLong deathsDayStat
		= new StatDayDataLong();

	private Map<AgeGroup25, StatDayDataLong> deathsDayStatByAge;
	private StatDayDataLong deathsDayStatAgeUnknown
		= new StatDayDataLong();

	public StatDay()
	{
	}

	// ========================================================================
	// = Deaths ===============================================================
	// ========================================================================
	
	public void addDeaths( long count )
	{
		addDeaths( count, null );
	}

	public void addDeaths( long count, AgeGroup25 age )
	{
		deathsDayStat.setMin( count );
		deathsDayStat.addToAverageValues( count );
		deathsDayStat.setMax( count );

		StatDayDataLong data;
		if( age != null )
		{
			data = deathsDayStatByAge.get( age );
			if( data == null )
			{
				data = new StatDayDataLong();
				deathsDayStatByAge.put( age, data );
			}

		}
		else
		{
			data = deathsDayStatAgeUnknown;
		}

		data.setMin(count);
		data.addToAverageValues(count);
		data.setMax(count);
	}

	// ------------------------------------------------------------------------

	public long getMinDeaths()
	{
		return deathsDayStat.getMin();
	}

	public double getAverageDeaths()
	{
		return deathsDayStat.getAvg();
	}

	public long getMaxDeaths()
	{
		return deathsDayStat.getMax();
	}

	// ------------------------------------------------------------------------

	public long getMinDeaths( AgeGroup25 age )
	{
		if( age == null )
		{
			return deathsDayStatAgeUnknown.getMin();
		}

		StatDayDataLong data = deathsDayStatByAge.get( age );
		if( data == null )
		{
			return 0;
		}

		return data.getMin();
	}

	public double getAverageDeaths( AgeGroup25 age )
	{
		if( age == null )
		{
			return deathsDayStatAgeUnknown.getAvg();
		}

		StatDayDataLong data = deathsDayStatByAge.get( age );
		if( data == null )
		{
			return 0.0d;
		}

		return data.getAvg();
	}

	public long getMaxDeaths( AgeGroup25 age )
	{
		if( age == null )
		{
			return deathsDayStatAgeUnknown.getMax();
		}

		StatDayDataLong data = deathsDayStatByAge.get( age );
		if( data == null )
		{
			return 0;
		}

		return data.getMax();
	}
}
