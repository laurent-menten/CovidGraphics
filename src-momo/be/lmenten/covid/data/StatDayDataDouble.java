package be.lmenten.covid.data;

import java.util.ArrayList;
import java.util.List;

public class StatDayDataDouble
{
	private List<Double> averageValues = new ArrayList<>();

	private double min;
	private double avg;
	private double max;

	// ========================================================================
	// =
	// ========================================================================

	public StatDayDataDouble()
	{
		this( 0.0d, 0.0d );
	}

	public StatDayDataDouble( double min, double max )
	{
		this.min = min;
		this.max = max;

		this.avg = (min + max) / 2.0d;
	}

	// ------------------------------------------------------------------------
	// - 
	// ------------------------------------------------------------------------

	public void setMin( long min )
	{
		if( min < this.min )
		{
			this.min = min;
		}
	}

	public void setMax( long max )
	{
		if( max > this.max )
		{
			this.max = max;
		}
	}

	// ------------------------------------------------------------------------

	public void addToAverageValues( double value )
	{
		averageValues.add( value );
	}

	public void resetAverageValues()
	{
		averageValues.clear();
	}

	public void computeAverage()
	{
		avg = averageValues.stream()
				.mapToDouble( val -> val )
					.average()
					.orElse(0.0)
				;
	}

	// ------------------------------------------------------------------------
	// - 
	// ------------------------------------------------------------------------

	public double getMin()
	{
		return min;
	}

	public double getAvg()
	{
		return avg;
	}

	public double getMax()
	{
		return max;
	}
}
